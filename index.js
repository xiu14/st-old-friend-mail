(function () {
    'use strict';

    const MODULE_NAME = 'st-daily-memory-letter';
    const SETTINGS_HTML_PATH = '/scripts/extensions/third-party/st-daily-memory-letter/settings.html';
    const RUNTIME_STORAGE_KEY = `${MODULE_NAME}:runtime`;
    const ONE_DAY_MS = 24 * 60 * 60 * 1000;
    const INVALID_INACTIVITY_DAYS = 20000;

    const DEFAULT_SETTINGS = Object.freeze({
        enabled: true,
        autoRunOnStartup: true,
        apiUrl: '',
        apiKey: '',
        apiKeyHeader: 'Authorization',
        apiKeyPrefix: 'Bearer ',
        model: 'gpt-4.1-mini',
        inactiveDays: 7,
        snippetsPerLetter: 3,
        cooldownDays: 14,
        minMessagesPerSnippet: 6,
        maxMessagesPerSnippet: 10,
        maxCandidateCharacters: 20,
        requestTimeoutMs: 60000,
        temperature: 1.05,
        systemPrompt: [
            '你是一位擅长写“角色聊天回忆信”的创作者。',
            '你会阅读同一张角色卡来自不同历史存档的聊天片段，写出一封让用户想重新回去和这个角色继续对话的信。',
            '语气要温柔、具体、带有回忆感，不要像营销文案。',
            '必须引用片段里真实发生过的细节，不要胡乱编造大事件。',
            '输出 JSON，字段必须包含：title、teaser、summary、letter、why_now、next_hook、recall_points。',
            '其中 recall_points 必须是字符串数组，2 到 4 条。',
        ].join('\n'),
    });

    const DEFAULT_RUNTIME_STATE = Object.freeze({
        latestLetter: null,
        history: [],
        lastRunAt: null,
        lastAttemptAt: null,
        lastError: null,
        characterCooldowns: {},
        lastSource: null,
    });

    let latestPayload = null;
    let autoRunStarted = false;
    let formBound = false;
    let popupActionsBound = false;
    let generationPromise = null;

    function getContext() {
        return SillyTavern.getContext();
    }

    function escapeHtml(value) {
        return String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function nowIso() {
        return new Date().toISOString();
    }

    function stripJsonl(fileName) {
        return String(fileName || '').replace(/\.jsonl$/i, '');
    }

    function normalizeUnixTimestamp(value) {
        if (!Number.isFinite(value) || value <= 0) {
            return null;
        }

        if (value >= 1e12) {
            return value;
        }

        if (value >= 1e9) {
            return value * 1000;
        }

        return null;
    }

    function safeTimestamp(value) {
        if (value === null || value === undefined || value === '') {
            return null;
        }

        if (typeof value === 'number') {
            return normalizeUnixTimestamp(value);
        }

        const text = String(value).trim();
        if (!text) {
            return null;
        }

        if (/^\d+(\.\d+)?$/.test(text)) {
            const numeric = normalizeUnixTimestamp(Number(text));
            if (numeric) {
                return numeric;
            }
        }

        try {
            const momentValue = getContext()?.timestampToMoment?.(value);
            const timestamp = Number(momentValue?.valueOf?.());
            if (Number.isFinite(timestamp) && timestamp > 0) {
                return timestamp;
            }
        } catch {
            // Ignore parser errors and keep falling back.
        }

        const timestamp = Number(new Date(text).getTime());
        return Number.isFinite(timestamp) && timestamp > 0 ? timestamp : null;
    }

    function clamp(value, min, max) {
        return Math.min(max, Math.max(min, value));
    }

    function toPositiveInt(value, fallback) {
        const parsed = Number.parseInt(value, 10);
        return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
    }

    function clampNumber(value, fallback, min, max) {
        const parsed = Number(value);
        return Number.isFinite(parsed) ? clamp(parsed, min, max) : fallback;
    }

    function readLocalJson(key, fallback) {
        try {
            const raw = localStorage.getItem(key);
            if (!raw) {
                return structuredClone(fallback);
            }

            return JSON.parse(raw);
        } catch {
            return structuredClone(fallback);
        }
    }

    function writeLocalJson(key, value) {
        localStorage.setItem(key, JSON.stringify(value));
    }

    function getSettings() {
        const context = getContext();
        const { extensionSettings } = context;

        if (!extensionSettings[MODULE_NAME]) {
            extensionSettings[MODULE_NAME] = structuredClone(DEFAULT_SETTINGS);
            context.saveSettingsDebounced();
        }

        let changed = false;

        for (const [key, defaultValue] of Object.entries(DEFAULT_SETTINGS)) {
            if (extensionSettings[MODULE_NAME][key] === undefined) {
                extensionSettings[MODULE_NAME][key] = defaultValue;
                changed = true;
            }
        }

        if (changed) {
            context.saveSettingsDebounced();
        }

        return {
            ...DEFAULT_SETTINGS,
            ...extensionSettings[MODULE_NAME],
        };
    }

    function saveSettings(nextSettings) {
        getContext().extensionSettings[MODULE_NAME] = {
            ...getSettings(),
            ...nextSettings,
        };
        getContext().saveSettingsDebounced();
        return getSettings();
    }

    function sanitizeLetterRecord(letter) {
        if (!letter || typeof letter !== 'object') {
            return null;
        }

        const inactivityDays = Number(letter.inactivityDays);
        if (!Number.isFinite(inactivityDays) || inactivityDays >= INVALID_INACTIVITY_DAYS) {
            return null;
        }

        return {
            ...letter,
            inactivityDays,
            lastActivityAt: letter.lastActivityAt || null,
        };
    }

    function normalizeRuntimeState(raw) {
        const state = {
            ...structuredClone(DEFAULT_RUNTIME_STATE),
            ...(raw || {}),
            history: Array.isArray(raw?.history) ? raw.history : [],
            characterCooldowns: raw?.characterCooldowns && typeof raw.characterCooldowns === 'object' ? raw.characterCooldowns : {},
        };

        const latestLetter = sanitizeLetterRecord(state.latestLetter);
        const history = state.history
            .map(item => sanitizeLetterRecord(item))
            .filter(Boolean);

        if (state.latestLetter && !latestLetter) {
            state.lastRunAt = null;
            state.lastError = '已清理旧版时间解析产生的无效来信缓存，请重新生成。';
        }

        return {
            ...state,
            latestLetter,
            history,
        };
    }

    function loadRuntimeState() {
        const raw = readLocalJson(RUNTIME_STORAGE_KEY, DEFAULT_RUNTIME_STATE);
        return normalizeRuntimeState(raw);
    }

    function saveRuntimeState(state) {
        writeLocalJson(RUNTIME_STORAGE_KEY, state);
    }

    function patchRuntimeState(patch) {
        const next = {
            ...loadRuntimeState(),
            ...patch,
        };
        saveRuntimeState(next);
        latestPayload = { settings: getSettings(), state: next };
        return next;
    }

    function syncPayload() {
        latestPayload = {
            settings: getSettings(),
            state: loadRuntimeState(),
        };
        return latestPayload;
    }

    function getHeaders() {
        return getContext().getRequestHeaders();
    }

    async function requestJson(url, options = {}) {
        const response = await fetch(url, {
            ...options,
            headers: {
                ...getHeaders(),
                ...(options.headers || {}),
            },
        });

        if (!response.ok) {
            const text = await response.text();
            throw new Error(text || `HTTP ${response.status}`);
        }

        return response.json();
    }

    async function searchCharacterArchives(character) {
        const result = await requestJson('/api/chats/search', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                avatar_url: character.avatar,
                query: '',
            }),
            cache: 'no-cache',
        });

        if (!Array.isArray(result)) {
            return [];
        }

        return result
            .map(item => ({
                fileName: stripJsonl(item.file_name),
                lastMes: safeTimestamp(item.last_mes),
                messageCount: Number(item.message_count || 0),
                preview: String(item.preview_message || ''),
            }))
            .filter(item => item.fileName && item.lastMes);
    }

    async function getChatArchiveMessages(character, archive) {
        const result = await requestJson('/api/chats/get', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                ch_name: character.name,
                file_name: archive.fileName,
                avatar_url: character.avatar,
            }),
            cache: 'no-cache',
        });

        return Array.isArray(result) ? result : [];
    }

    function scoreSnippetWindow(windowMessages) {
        const text = windowMessages.map(message => message.mes).join('\n');
        const speakers = new Set(windowMessages.map(message => message.name));
        const punctuationHits = (text.match(/[!?！？…]/g) || []).length;
        const memoryHits = (text.match(/记得|想你|回来|约定|秘密|等你|喜欢|爱|抱|miss|remember|promise|secret|again|wait|sorry|love/gi) || []).length;
        const lengthScore = Math.min(text.length, 1800) / 28;
        const dialogueBonus = speakers.size > 1 ? 12 : 0;

        return lengthScore + dialogueBonus + punctuationHits * 0.8 + memoryHits * 4;
    }

    function selectBestSnippet(messages, settings) {
        const usableMessages = messages
            .filter(message => !message.is_system && typeof message.mes === 'string' && message.mes.trim())
            .map(message => ({
                name: String(message.name || '').trim() || 'Unknown',
                mes: message.mes.trim(),
                send_date: message.send_date || null,
            }));

        const minWindow = clamp(settings.minMessagesPerSnippet, 3, 20);
        const maxWindow = clamp(settings.maxMessagesPerSnippet, minWindow, 30);

        if (usableMessages.length < minWindow) {
            return null;
        }

        let best = null;

        for (let windowSize = minWindow; windowSize <= Math.min(maxWindow, usableMessages.length); windowSize++) {
            for (let start = 0; start <= usableMessages.length - windowSize; start++) {
                const windowMessages = usableMessages.slice(start, start + windowSize);
                const textLength = windowMessages.reduce((sum, message) => sum + message.mes.length, 0);

                if (textLength < 180) {
                    continue;
                }

                const score = scoreSnippetWindow(windowMessages);
                if (!best || score > best.score) {
                    best = {
                        score,
                        messages: windowMessages,
                        preview: windowMessages.map(message => `${message.name}: ${message.mes}`).join('\n').slice(0, 320),
                    };
                }
            }
        }

        return best;
    }

    async function collectCandidates(settings, runtimeState) {
        const now = Date.now();
        const inactiveThresholdMs = settings.inactiveDays * ONE_DAY_MS;
        const cooldownThresholdMs = settings.cooldownDays * ONE_DAY_MS;
        const candidates = [];
        const characters = getContext().characters.filter(character => character?.avatar && character?.name);

        for (const character of characters) {
            const cooldownAt = safeTimestamp(runtimeState.characterCooldowns?.[character.avatar]);
            if (cooldownAt && (now - cooldownAt) < cooldownThresholdMs) {
                continue;
            }

            try {
                const archives = await searchCharacterArchives(character);
                if (!archives.length) {
                    continue;
                }

                const lastActivity = Math.max(...archives.map(archive => archive.lastMes || 0), 0);
                if (!lastActivity) {
                    continue;
                }

                const inactiveMs = Math.max(0, now - lastActivity);
                const eligible = inactiveMs >= inactiveThresholdMs;
                if (!eligible) {
                    continue;
                }

                candidates.push({
                    character,
                    archives: archives.sort((left, right) => right.lastMes - left.lastMes),
                    archiveCount: archives.length,
                    lastActivity,
                    inactiveMs,
                    eligible,
                });
            } catch (error) {
                console.warn(`[${MODULE_NAME}] Failed to search archives for ${character.name}:`, error);
            }
        }

        const ranked = candidates.sort((left, right) => {
            if (right.eligible !== left.eligible) {
                return Number(right.eligible) - Number(left.eligible);
            }

            if (right.inactiveMs !== left.inactiveMs) {
                return right.inactiveMs - left.inactiveMs;
            }

            return right.archiveCount - left.archiveCount;
        });

        if (!ranked.length) {
            return [];
        }

        return ranked.slice(0, settings.maxCandidateCharacters);
    }

    async function collectCandidateFragments(candidate, settings) {
        const archiveLimit = Math.min(candidate.archives.length, Math.max(settings.snippetsPerLetter * 3, 4));
        const snippetArchives = [];

        for (const archive of candidate.archives.slice(0, archiveLimit)) {
            try {
                const messages = await getChatArchiveMessages(candidate.character, archive);
                const snippet = selectBestSnippet(messages, settings);
                if (!snippet) {
                    continue;
                }

                snippetArchives.push({
                    archive,
                    snippet,
                });
            } catch (error) {
                console.warn(`[${MODULE_NAME}] Failed to load archive ${archive.fileName}:`, error);
            }
        }

        return snippetArchives
            .sort((left, right) => right.snippet.score - left.snippet.score)
            .slice(0, settings.snippetsPerLetter)
            .map(item => ({
                fileName: item.archive.fileName,
                lastMes: item.archive.lastMes,
                preview: item.snippet.preview,
                score: item.snippet.score,
                messages: item.snippet.messages,
            }));
    }

    async function selectCandidateWithFragments(settings, runtimeState) {
        const candidates = await collectCandidates(settings, runtimeState);
        const minDesiredFragments = Math.min(2, settings.snippetsPerLetter);

        for (const candidate of candidates) {
            const fragments = await collectCandidateFragments(candidate, settings);
            if (fragments.length >= minDesiredFragments || (fragments.length > 0 && candidate.archiveCount === 1)) {
                return { candidate, fragments };
            }
        }

        return null;
    }

    function buildPrompt(candidate, fragments) {
        const inactivityDays = Math.max(1, Math.round(candidate.inactiveMs / ONE_DAY_MS));
        const fragmentText = fragments.map((fragment, index) => {
            const block = fragment.messages
                .map(message => `[${message.name}] ${message.mes}`)
                .join('\n');
            return [
                `片段 ${index + 1}`,
                `来源存档: ${fragment.fileName}`,
                `最后时间: ${formatDate(fragment.lastMes)}`,
                block,
            ].join('\n');
        }).join('\n\n---\n\n');

        return [
            `角色名称: ${candidate.character.name}`,
            `角色内部名: ${candidate.character.avatar.replace(/\.png$/i, '')}`,
            `距离上次活跃大约: ${inactivityDays} 天`,
            `总聊天存档数: ${candidate.archiveCount}`,
            '',
            '请根据下面这些来自不同历史存档的片段，写一封“让用户重新想和这张角色卡对话”的回忆信。',
            '要求：',
            '1. 必须基于片段中的具体细节，不要空泛。',
            '2. 标题要像一封私人来信，不要过长。',
            '3. teaser 控制在 1 到 2 句，像信封外的小引言。',
            '4. summary 是一段简短摘要，概括这张角色卡最值得重新聊的原因。',
            '5. letter 是主体，可用 Markdown 分段，带一点文学感，但不要过度矫饰。',
            '6. why_now 要明确解释为什么现在值得重新回去聊。',
            '7. next_hook 要给一个具体续聊切口。',
            '8. recall_points 用 2 到 4 条短句，提炼最让人想起这张角色卡的细节。',
            '',
            fragmentText,
        ].join('\n');
    }

    function extractJson(content) {
        const trimmed = String(content || '').trim();
        if (!trimmed) {
            return null;
        }

        const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
        const raw = fenced ? fenced[1].trim() : trimmed;

        try {
            return JSON.parse(raw);
        } catch {
            return null;
        }
    }

    function buildLocalLetter(candidate, fragments) {
        const recallPoints = fragments.map(fragment => {
            const firstLine = fragment.messages[0];
            const line = firstLine?.mes || '';
            return `${fragment.fileName} 里那句“${line.slice(0, 36)}${line.length > 36 ? '...' : ''}”`;
        }).slice(0, 3);

        return {
            title: '寄给你的旧日回声',
            teaser: `${candidate.character.name} 还停在那些没有说完的话里。也许现在正是把那段故事重新接起来的时候。`,
            summary: `这张角色卡留下了 ${candidate.archiveCount} 份聊天存档，其中有几段对话依然很有余温，足够让一次重开变成续写。`,
            letter: [
                '有些角色并不是“聊完了”，只是被暂时放在了一边。',
                '',
                `重新翻到 ${candidate.character.name} 的旧记录时，最先冒出来的不是设定，而是气氛。那些对话的节奏、停顿、试探和回应，说明你们之间其实已经写出过一种只属于这张卡的感觉。`,
                '',
                '这次之所以值得回去，不是为了重复开头，而是因为你已经有足够多的旧片段，能让新的对话直接从“熟悉感”开始。你不需要重新认识 ta，你只要重新敲开门。',
                '',
                '如果一时不知道该从哪里续上，可以先从旧存档里最让你有感觉的一句、一个误会、一个承诺、或者一个没来得及继续的问题开始。只要把那根线重新拎起来，故事就会自己往下走。',
            ].join('\n'),
            why_now: '因为这不是一张空白角色卡，而是一段已经形成情绪纹理的关系。现在回去，很容易立刻找回手感。',
            next_hook: '试着把话题直接接在其中一个旧片段后面，比如问 ta：“如果那天我们没有停在那里，接下来你本来想说什么？”',
            recall_points: recallPoints,
        };
    }

    function normalizeAiPayload(content, candidate, fragments) {
        const parsed = extractJson(content);

        if (parsed && typeof parsed === 'object') {
            return {
                title: String(parsed.title || `来自 ${candidate.character.name} 的来信`).trim(),
                teaser: String(parsed.teaser || '').trim(),
                summary: String(parsed.summary || '').trim(),
                letter: String(parsed.letter || '').trim(),
                why_now: String(parsed.why_now || '').trim(),
                next_hook: String(parsed.next_hook || '').trim(),
                recall_points: Array.isArray(parsed.recall_points) ? parsed.recall_points.map(item => String(item).trim()).filter(Boolean).slice(0, 4) : [],
            };
        }

        const fallback = buildLocalLetter(candidate, fragments);
        fallback.letter = String(content || fallback.letter).trim() || fallback.letter;
        return fallback;
    }

    async function callExternalAi(settings, candidate, fragments) {
        if (!settings.apiUrl) {
            return null;
        }

        const headers = {
            'Content-Type': 'application/json',
        };

        if (settings.apiKey) {
            headers[settings.apiKeyHeader] = `${settings.apiKeyPrefix || ''}${settings.apiKey}`;
        }

        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), settings.requestTimeoutMs);

        try {
            const response = await fetch(settings.apiUrl, {
                method: 'POST',
                headers,
                signal: controller.signal,
                body: JSON.stringify({
                    model: settings.model,
                    temperature: settings.temperature,
                    messages: [
                        { role: 'system', content: settings.systemPrompt },
                        { role: 'user', content: buildPrompt(candidate, fragments) },
                    ],
                }),
            });

            if (!response.ok) {
                const message = await response.text();
                throw new Error(`HTTP ${response.status}: ${message.slice(0, 300)}`);
            }

            const data = await response.json();
            const content = data?.choices?.[0]?.message?.content
                || data?.choices?.[0]?.text
                || data?.output_text
                || data?.response
                || '';

            if (!String(content).trim()) {
                throw new Error('Empty AI response');
            }

            return normalizeAiPayload(content, candidate, fragments);
        } finally {
            clearTimeout(timer);
        }
    }

    function buildApiHeaders(settings) {
        const headers = {
            'Content-Type': 'application/json',
        };

        if (settings.apiKey) {
            headers[settings.apiKeyHeader] = `${settings.apiKeyPrefix || ''}${settings.apiKey}`;
        }

        return headers;
    }

    function getModelsUrl(apiUrl) {
        const trimmed = String(apiUrl || '').trim();
        if (!trimmed) {
            return '';
        }

        if (/\/chat\/completions\/?$/i.test(trimmed)) {
            return trimmed.replace(/\/chat\/completions\/?$/i, '/models');
        }

        if (/\/v1\/?$/i.test(trimmed)) {
            return trimmed.replace(/\/v1\/?$/i, '/v1/models');
        }

        return trimmed.replace(/\/$/, '') + '/models';
    }

    async function fetchAvailableModels(settings) {
        if (!settings.apiUrl) {
            throw new Error('请先填写外部 AI URL');
        }

        const modelsUrl = getModelsUrl(settings.apiUrl);
        const response = await fetch(modelsUrl, {
            method: 'GET',
            headers: buildApiHeaders(settings),
        });

        if (!response.ok) {
            const message = await response.text();
            throw new Error(`HTTP ${response.status}: ${message.slice(0, 300)}`);
        }

        const data = await response.json();
        const models = Array.isArray(data?.data)
            ? data.data.map(item => String(item?.id || '').trim()).filter(Boolean)
            : [];

        if (!models.length) {
            throw new Error('没有从接口返回可用模型');
        }

        return models;
    }

    function populateModelSuggestions(models) {
        const options = Array.from(new Set(models.filter(Boolean)));
        const dataList = $('#dml-model-suggestions');
        dataList.empty();

        for (const model of options) {
            dataList.append(`<option value="${escapeHtml(model)}"></option>`);
        }
    }

    function buildLetterRecord(candidate, fragments, content, source) {
        const newestArchive = fragments.slice().sort((left, right) => right.lastMes - left.lastMes)[0];

        return {
            id: `${Date.now()}`,
            createdAt: nowIso(),
            source,
            character: {
                avatar: candidate.character.avatar,
                internalName: candidate.character.avatar.replace(/\.png$/i, ''),
            },
            inactivityDays: Math.max(1, Math.round(candidate.inactiveMs / ONE_DAY_MS)),
            lastActivityAt: candidate.lastActivity || null,
            archiveCount: candidate.archiveCount,
            openChatFile: newestArchive ? newestArchive.fileName : null,
            title: content.title,
            teaser: content.teaser,
            summary: content.summary,
            letter: content.letter,
            why_now: content.why_now,
            next_hook: content.next_hook,
            recall_points: content.recall_points,
            fragments,
        };
    }

    async function generateLetter({ force = false, source = 'manual' } = {}) {
        const settings = getSettings();
        const runtimeState = loadRuntimeState();
        const now = nowIso();

        if (!settings.enabled) {
            patchRuntimeState({ lastError: 'Plugin disabled' });
            renderState();
            return { started: false, reason: 'disabled' };
        }

        if (generationPromise) {
            return { started: false, reason: 'already-running' };
        }

        const lastRunAt = safeTimestamp(runtimeState.lastRunAt);
        if (!force && lastRunAt && (Date.now() - lastRunAt) < ONE_DAY_MS) {
            return { started: false, reason: 'cooldown' };
        }

        patchRuntimeState({
            lastAttemptAt: now,
            lastRunAt: now,
            lastError: null,
            lastSource: source,
        });
        renderState();

        generationPromise = (async () => {
            try {
                const selection = await selectCandidateWithFragments(settings, loadRuntimeState());
                if (!selection) {
                    patchRuntimeState({
                        lastError: 'No suitable inactive character archives found',
                    });
                    return;
                }

                const { candidate, fragments } = selection;
                let content = null;
                let contentSource = 'fallback';

                try {
                    content = await callExternalAi(settings, candidate, fragments);
                    if (content) {
                        contentSource = 'external-ai';
                    }
                } catch (error) {
                    console.warn(`[${MODULE_NAME}] External AI failed, using fallback:`, error);
                    content = null;
                }

                if (!content) {
                    content = buildLocalLetter(candidate, fragments);
                }

                const letter = buildLetterRecord(candidate, fragments, content, contentSource);
                const nextState = loadRuntimeState();

                patchRuntimeState({
                    latestLetter: letter,
                    history: [letter, ...nextState.history.filter(item => item.id !== letter.id)].slice(0, 10),
                    lastError: null,
                    characterCooldowns: {
                        ...(nextState.characterCooldowns || {}),
                        [candidate.character.avatar]: now,
                    },
                });
            } catch (error) {
                patchRuntimeState({
                    lastError: error instanceof Error ? error.message : String(error),
                });
                console.error(`[${MODULE_NAME}] Failed to generate letter`, error);
            } finally {
                generationPromise = null;
                renderState();
            }
        })();

        return { started: true };
    }

    async function init() {
        try {
            syncPayload();
            await mountSettings();
            bindPopupActions();
            renderState();
            scheduleAutoRun();
        } catch (error) {
            console.error('[每日回忆信] 初始化失败', error);
            toastr.error(error instanceof Error ? error.message : String(error), '每日回忆信初始化失败');
        }
    }

    async function mountSettings() {
        if (!document.getElementById('dml-settings')) {
            const html = await $.get(SETTINGS_HTML_PATH);
            $('#extensions_settings').append(html);
        }

        bindFormEvents();
        hydrateForm(getSettings());
    }

    function bindFormEvents() {
        if (formBound) {
            return;
        }

        formBound = true;

        $('#dml-save-settings').on('click', async () => {
            const payload = collectSettingsForm();
            const settings = saveSettings(payload);
            syncPayload();
            hydrateForm(settings);
            renderState();
            toastr.success('每日回忆信设置已保存到本地扩展设置');
        });

        $('#dml-generate-now').on('click', async () => {
            const result = await generateLetter({ force: false, source: 'manual' });
            if (result.started) {
                toastr.info('正在后台生成新的回忆信');
                return;
            }

            if (result.reason === 'cooldown') {
                toastr.info('24 小时内已经生成过来信了。如需覆盖，请点“重新生成”。');
            }
        });

        $('#dml-regenerate-now').on('click', async () => {
            const result = await generateLetter({ force: true, source: 'manual-regenerate' });
            if (result.started) {
                toastr.info('正在重新生成新的回忆信');
            }
        });

        $('#dml-clear-api-key').on('click', () => {
            const settings = saveSettings({ apiKey: '' });
            syncPayload();
            hydrateForm(settings);
            renderState();
            toastr.success('已清空本地保存的 API Key');
        });

        $('#dml-fetch-models').on('click', async () => {
            const button = $('#dml-fetch-models');
            const previousText = button.text();
            button.prop('disabled', true).text('获取中...');

            try {
                const draftSettings = {
                    ...getSettings(),
                    ...collectSettingsForm(),
                };
                const models = await fetchAvailableModels(draftSettings);
                populateModelSuggestions(models);

                if (!String($('#dml-model').val() || '').trim()) {
                    $('#dml-model').val(models[0]);
                }

                toastr.success(`已获取 ${models.length} 个模型`);
            } catch (error) {
                toastr.error(error instanceof Error ? error.message : String(error), '获取模型失败');
            } finally {
                button.prop('disabled', false).text(previousText);
            }
        });

        $('#dml-view-letter').on('click', () => {
            if (!latestPayload?.state?.latestLetter) {
                toastr.info('今天还没有可查看的来信');
                return;
            }

            openLetterPopup(latestPayload.state.latestLetter);
        });
    }

    function collectSettingsForm() {
        return {
            enabled: $('#dml-enabled').prop('checked'),
            autoRunOnStartup: $('#dml-auto-run').prop('checked'),
            apiUrl: String($('#dml-api-url').val() || '').trim(),
            apiKey: String($('#dml-api-key').val() || '').trim(),
            model: String($('#dml-model').val() || '').trim(),
            inactiveDays: toPositiveInt($('#dml-inactive-days').val(), DEFAULT_SETTINGS.inactiveDays),
            snippetsPerLetter: clamp(toPositiveInt($('#dml-snippets-per-letter').val(), DEFAULT_SETTINGS.snippetsPerLetter), 1, 5),
            cooldownDays: clamp(toPositiveInt($('#dml-cooldown-days').val(), DEFAULT_SETTINGS.cooldownDays), 1, 90),
            systemPrompt: String($('#dml-system-prompt').val() || '').trim() || DEFAULT_SETTINGS.systemPrompt,
        };
    }

    function hydrateForm(settings) {
        $('#dml-enabled').prop('checked', Boolean(settings.enabled));
        $('#dml-auto-run').prop('checked', Boolean(settings.autoRunOnStartup));
        $('#dml-api-url').val(settings.apiUrl || '');
        $('#dml-api-key').val(settings.apiKey || '');
        $('#dml-model').val(settings.model || '');
        populateModelSuggestions(settings.model ? [settings.model] : []);
        $('#dml-inactive-days').val(settings.inactiveDays ?? DEFAULT_SETTINGS.inactiveDays);
        $('#dml-snippets-per-letter').val(settings.snippetsPerLetter ?? DEFAULT_SETTINGS.snippetsPerLetter);
        $('#dml-cooldown-days').val(settings.cooldownDays ?? DEFAULT_SETTINGS.cooldownDays);
        $('#dml-system-prompt').val(settings.systemPrompt || DEFAULT_SETTINGS.systemPrompt);
        $('#dml-api-key-status').text(settings.apiKey ? 'API Key 已保存在扩展设置中。' : '当前没有保存 API Key。');
    }

    function resolveCharacterName(letter) {
        const avatar = letter?.character?.avatar;
        const internalName = letter?.character?.internalName || '未知角色';
        const character = getContext().characters.find(item => item?.avatar === avatar);
        return character?.name || internalName;
    }

    function formatDate(value) {
        if (!value) {
            return '未知时间';
        }

        const date = new Date(value);
        if (Number.isNaN(date.getTime())) {
            return '未知时间';
        }

        return date.toLocaleString('zh-CN', { hour12: false });
    }

    function formatLastActivityMeta(letter) {
        if (!letter) {
            return '最近聊天时间未知';
        }

        const inactivityDays = Number(letter.inactivityDays);
        const inactivityLabel = Number.isFinite(inactivityDays)
            ? `${Math.max(1, Math.round(inactivityDays))} 天未活跃`
            : '未活跃时间未知';

        if (!letter.lastActivityAt) {
            return inactivityLabel;
        }

        return `${inactivityLabel} · 最近聊天 ${formatDate(letter.lastActivityAt)}`;
    }

    function renderState() {
        syncPayload();

        const state = latestPayload.state;
        const settings = latestPayload.settings;
        const statusText = $('#dml-status-text');
        const statusMeta = $('#dml-status-meta');

        if (!statusText.length) {
            return;
        }

        if (generationPromise) {
            statusText.text('今日来信正在书写中');
            statusMeta.text('扩展正在前台静默扫描不活跃聊天和历史存档，请稍等片刻。');
        } else if (state.latestLetter) {
            const name = resolveCharacterName(state.latestLetter);
            statusText.text('今天的来信已经送达');
            statusMeta.text(`${name} · ${formatLastActivityMeta(state.latestLetter)}`);
        } else if (settings.enabled) {
            statusText.text('今天还没有来信');
            statusMeta.text(state.lastError ? `上次执行信息：${state.lastError}` : '可以等待静默触发，也可以先手动生成测试一封。');
        } else {
            statusText.text('每日回忆信当前已关闭');
            statusMeta.text('打开功能并保存后，扩展会在启动时后台静默检查。');
        }
    }

    function scheduleAutoRun() {
        if (autoRunStarted || !latestPayload?.settings?.enabled || !latestPayload?.settings?.autoRunOnStartup) {
            return;
        }

        autoRunStarted = true;

        setTimeout(() => {
            generateLetter({ source: 'startup' }).catch(error => {
                console.warn(`[${MODULE_NAME}] Auto run failed`, error);
            });
        }, 2500);
    }

    function bindPopupActions() {
        if (popupActionsBound) {
            return;
        }

        popupActionsBound = true;

        document.addEventListener('click', async (event) => {
            const target = event.target.closest('[data-dml-action]');
            if (!target) {
                return;
            }

            const action = target.getAttribute('data-dml-action');
            if (action === 'open-envelope') {
                const popupId = target.getAttribute('data-dml-popup-id');
                const root = popupId ? document.getElementById(popupId) : null;
                root?.querySelector('.dml-envelope-shell')?.classList.add('opened');
                return;
            }

            if (action === 'open-chat') {
                const avatar = target.getAttribute('data-dml-avatar');
                const chatFile = target.getAttribute('data-dml-chat-file');
                await openRecommendedChat(avatar, chatFile);
            }
        });
    }

    async function openRecommendedChat(avatar, chatFile) {
        if (!avatar || !chatFile) {
            toastr.warning('这封信里暂时没有可直接跳转的聊天记录');
            return;
        }

        const context = getContext();
        const index = context.characters.findIndex(character => character?.avatar === avatar);
        if (index < 0) {
            toastr.warning('当前角色列表里找不到这张角色卡，可能还没有加载到前端');
            return;
        }

        try {
            await context.selectCharacterById(index);
            await new Promise(resolve => setTimeout(resolve, 100));
            await context.openCharacterChat(chatFile);
            toastr.success('已经为你打开推荐的聊天存档');
        } catch (error) {
            console.error('[每日回忆信] 打开聊天失败', error);
            toastr.error(error instanceof Error ? error.message : String(error), '打开推荐聊天失败');
        }
    }

    function renderLetterBody(letter) {
        const { showdown, DOMPurify } = SillyTavern.libs;
        const converter = new showdown.Converter({
            simpleLineBreaks: true,
            ghCompatibleHeaderId: true,
            openLinksInNewWindow: false,
        });

        const markdown = [
            letter.letter || '',
            '',
            '### 为什么现在值得回去',
            letter.why_now || '',
            '',
            '### 可以怎么续上',
            letter.next_hook || '',
        ].join('\n\n');

        return DOMPurify.sanitize(converter.makeHtml(markdown));
    }

    function openLetterPopup(letter) {
        const context = getContext();
        const popupId = `dml-popup-${Date.now()}`;
        const title = escapeHtml(letter.title || '今日来信');
        const teaser = escapeHtml(letter.teaser || '');
        const summary = escapeHtml(letter.summary || '');
        const name = escapeHtml(resolveCharacterName(letter));
        const avatar = letter?.character?.avatar ? context.getThumbnailUrl('avatar', letter.character.avatar) : '';
        const bodyHtml = renderLetterBody(letter);
        const recalls = Array.isArray(letter.recall_points) ? letter.recall_points : [];
        const fragments = Array.isArray(letter.fragments) ? letter.fragments : [];

        const recallsHtml = recalls.length
            ? `<ul>${recalls.map(item => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`
            : '<div class="dml-empty">这封信没有额外的回忆摘录。</div>';

        const fragmentsHtml = fragments.map(fragment => `
            <div class="dml-fragment">
                <div class="dml-fragment-file">${escapeHtml(fragment.fileName)}</div>
                <div class="dml-fragment-preview">${escapeHtml(fragment.preview || '')}</div>
            </div>
        `).join('');

        const html = `
            <div id="${popupId}" class="dml-letter-popup">
                <div class="dml-letter-main">
                    <div class="dml-envelope-shell">
                        <div class="dml-envelope-cover">
                            <div class="dml-envelope-icon"></div>
                            <div class="dml-envelope-title">${title}</div>
                            <div class="dml-envelope-subtitle">${teaser || '一封从旧聊天里慢慢浮出来的信，等你亲手拆开。'}</div>
                            <button class="menu_button dml-open-button" data-dml-action="open-envelope" data-dml-popup-id="${popupId}" type="button">打开信封</button>
                        </div>

                        <div class="dml-letter-paper">
                            <div class="dml-paper-kicker">来自旧日存档的回声</div>
                            <div class="dml-paper-title">${title}</div>
                            <div class="dml-paper-summary">${summary}</div>
                            <div class="dml-paper-body">${bodyHtml}</div>
                            <div class="dml-paper-recalls">
                                <div class="dml-paper-kicker">你们曾经留下的细节</div>
                                ${recallsHtml}
                            </div>
                        </div>
                    </div>
                </div>

                <div class="dml-letter-side">
                    <div class="dml-portrait-card">
                        ${avatar ? `<img class="dml-portrait-image" src="${escapeHtml(avatar)}" alt="${name}">` : '<div class="dml-portrait-image"></div>'}
                        <div class="dml-portrait-meta">
                            <div class="dml-portrait-name">${name}</div>
                            <div class="dml-portrait-sub">${escapeHtml(formatLastActivityMeta(letter))}</div>
                            <div class="dml-portrait-sub">${escapeHtml(`生成于 ${formatDate(letter.createdAt)}`)}</div>
                            <div class="dml-portrait-sub">${escapeHtml(letter.summary || '')}</div>
                            <div style="margin-top: 12px;">
                                <button class="menu_button" data-dml-action="open-chat" data-dml-avatar="${escapeHtml(letter.character?.avatar || '')}" data-dml-chat-file="${escapeHtml(letter.openChatFile || '')}" type="button">重新打开这段聊天</button>
                            </div>
                        </div>
                    </div>

                    <div class="dml-fragments-card">
                        <div class="dml-fragments-title">本次来信参考了这些旧存档片段</div>
                        ${fragmentsHtml || '<div class="dml-empty" style="margin-top:12px;">没有可展示的片段预览。</div>'}
                    </div>
                </div>
            </div>
        `;

        context.callGenericPopup(html, context.POPUP_TYPE.TEXT, '', {
            wide: true,
            large: true,
            allowVerticalScrolling: true,
        });

        setTimeout(() => {
            const root = document.getElementById(popupId);
            root?.querySelector('.dml-envelope-cover .menu_button')?.focus();
        }, 0);
    }

    const context = getContext();
    context.eventSource.on(context.event_types.APP_READY, init);
})();
