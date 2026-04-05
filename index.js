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
        useLocalGeneration: false,
        useInCharacterMode: false,
        apiUrl: '',
        apiKey: '',
        apiKeyHeader: 'Authorization',
        apiKeyPrefix: 'Bearer ',
        model: 'gpt-4.1-mini',
        inactiveDays: 7,
        snippetsPerLetter: 3,
        cooldownDays: 14,
        contentTagName: 'content',
        minMessagesPerSnippet: 6,
        maxMessagesPerSnippet: 10,
        maxCandidateCharacters: 20,
        requestTimeoutMs: 120000,
        temperature: 1.05,
        analysisSystemPrompt: [
            '你是一位擅长写“故人来信”的创作者。',
            '你会阅读同一张角色卡来自不同历史存档的聊天片段，写出一封让用户想重新回去和这个角色继续对话的信。',
            '语气要温柔、具体、带有回忆感，不要像营销文案。',
            '必须引用片段里真实发生过的细节，不要胡乱编造大事件。',
            '输出 JSON，字段必须包含：title、teaser、summary、letter、why_now、next_hook、recall_points。',
            '其中 recall_points 必须是字符串数组，2 到 4 条。',
        ].join('\n'),
        inCharacterSystemPrompt: [
            '你现在要扮演这张角色卡本人，给用户写一封第一人称来信。',
            '你会同时参考角色卡设定和来自不同历史存档的聊天片段。',
            '整封信要像角色亲自写给用户，而不是旁白分析或作者说明。',
            '必须保持角色的语气、价值观、关系感和世界观，不要跳出角色。',
            '必须引用聊天片段里真实发生过的细节，不要凭空编造重大事件。',
            '输出 JSON，字段必须包含：title、teaser、summary、letter、why_now、next_hook、recall_points。',
            '其中 teaser、summary、letter、why_now、next_hook 应当全部使用角色第一人称口吻。',
            'teaser 和 summary 不能写成“这张角色卡值得聊”的旁白分析句，必须像角色正在对用户说话。',
            'recall_points 必须是字符串数组，2 到 4 条。',
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

    function normalizeContentTagName(value) {
        const text = String(value || '')
            .trim()
            .replace(/^<\s*/, '')
            .replace(/\s*>$/, '')
            .replace(/^\/+|\/+$/g, '')
            .trim();

        if (!text) {
            return '';
        }

        const match = text.match(/^[A-Za-z][\w:-]*/);
        return match ? match[0] : '';
    }

    function escapeRegExp(value) {
        return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    function shouldUseLocalGeneration(settings) {
        return Boolean(settings?.useLocalGeneration);
    }

    function canGenerateWithApi(settings) {
        return Boolean(String(settings?.apiUrl || '').trim());
    }

    function isInCharacterMode(settingsOrLetter) {
        return Boolean(settingsOrLetter?.useInCharacterMode || settingsOrLetter?.mode === 'in_character');
    }

    function getActiveSystemPrompt(settings) {
        return isInCharacterMode(settings)
            ? (settings.inCharacterSystemPrompt || DEFAULT_SETTINGS.inCharacterSystemPrompt)
            : (settings.analysisSystemPrompt || DEFAULT_SETTINGS.analysisSystemPrompt);
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

        if (extensionSettings[MODULE_NAME].systemPrompt && extensionSettings[MODULE_NAME].analysisSystemPrompt === DEFAULT_SETTINGS.analysisSystemPrompt) {
            extensionSettings[MODULE_NAME].analysisSystemPrompt = extensionSettings[MODULE_NAME].systemPrompt;
            changed = true;
        }

        // Migrate the old hidden timeout default to the new 120s baseline.
        if (Number(extensionSettings[MODULE_NAME].requestTimeoutMs) === 60000) {
            extensionSettings[MODULE_NAME].requestTimeoutMs = DEFAULT_SETTINGS.requestTimeoutMs;
            changed = true;
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
            state.lastError = '已清理旧版时间解析产生的无效故人来信缓存，请重新生成。';
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

    function normalizeSnippetText(text) {
        return String(text || '')
            .toLowerCase()
            .replace(/[^\p{L}\p{N}\s]/gu, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function getSnippetSignature(snippet) {
        return normalizeSnippetText(
            Array.isArray(snippet?.messages)
                ? snippet.messages.map(message => `${message.name}: ${message.mes}`).join('\n')
                : (snippet?.preview || ''),
        );
    }

    function getTokenOverlapRatio(leftText, rightText) {
        const leftTokens = new Set(normalizeSnippetText(leftText).split(' ').filter(token => token.length > 1));
        const rightTokens = new Set(normalizeSnippetText(rightText).split(' ').filter(token => token.length > 1));

        if (!leftTokens.size || !rightTokens.size) {
            return 0;
        }

        let overlap = 0;
        for (const token of leftTokens) {
            if (rightTokens.has(token)) {
                overlap += 1;
            }
        }

        return overlap / Math.min(leftTokens.size, rightTokens.size);
    }

    function areSnippetsTooSimilar(leftSnippet, rightSnippet) {
        const leftSignature = getSnippetSignature(leftSnippet);
        const rightSignature = getSnippetSignature(rightSnippet);

        if (!leftSignature || !rightSignature) {
            return false;
        }

        if (leftSignature === rightSignature) {
            return true;
        }

        if (leftSignature.length > 80 && rightSignature.length > 80) {
            if (leftSignature.includes(rightSignature) || rightSignature.includes(leftSignature)) {
                return true;
            }
        }

        return getTokenOverlapRatio(leftSignature, rightSignature) >= 0.82;
    }

    function stripMarkup(text) {
        return String(text || '')
            .replace(/<[^>]*>/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function stripReasoningBlocks(text) {
        let cleaned = String(text || '');
        const patterns = [
            /<\s*(think|thinking|cot|reasoning|analysis|thought|inner_monologue)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi,
            /&lt;\s*(think|thinking|cot|reasoning|analysis|thought|inner_monologue)\b[^&]*&gt;[\s\S]*?&lt;\s*\/\s*\1\s*&gt;/gi,
            /\[\s*(think|thinking|cot|reasoning|analysis|thought|inner_monologue)\s*\][\s\S]*?\[\s*\/\s*\1\s*\]/gi,
        ];

        for (const pattern of patterns) {
            cleaned = cleaned.replace(pattern, ' ');
        }

        return cleaned;
    }

    function extractTaggedBlocks(raw, tagName) {
        const escapedTag = escapeRegExp(tagName);
        const patterns = [
            new RegExp(`<${escapedTag}\\b[^>]*>([\\s\\S]*?)<\\/${escapedTag}>`, 'gi'),
            new RegExp(`&lt;${escapedTag}\\b[^&]*&gt;([\\s\\S]*?)&lt;\\/${escapedTag}&gt;`, 'gi'),
            new RegExp(`\\[${escapedTag}\\]([\\s\\S]*?)\\[\\/${escapedTag}\\]`, 'gi'),
        ];

        const blocks = [];
        for (const pattern of patterns) {
            for (const match of raw.matchAll(pattern)) {
                blocks.push(String(match[1] || ''));
            }
        }

        return blocks;
    }

    function scoreTaggedContentBlock(text) {
        const normalized = String(text || '').trim();
        if (!normalized) {
            return Number.NEGATIVE_INFINITY;
        }

        let score = Math.min(normalized.length, 1200) / 40;
        const rewardPatterns = [
            /[。！？?!…]/g,
            /[“”"'「」『』]/g,
            /\b(我|你|他|她|它|我们|他们|然后|只是|因为|如果|可是|仍然|已经)\b/g,
        ];
        const penaltyPatterns = [
            /(状态栏|状态|摘要|总结|概括|提要|大纲|要点|禁用词|标签|数值|好感|数值变化|属性|面板|注意事项|要求|思考|分析|推理|cot|reasoning)/gi,
            /(^|\n)\s*[-*•]\s*/g,
            /(^|\n)\s*\d+\.\s*/g,
            /[:：]\s*/g,
            /[【\[][^【\]]{0,18}[】\]]/g,
        ];

        for (const pattern of rewardPatterns) {
            score += (normalized.match(pattern) || []).length * 2.5;
        }

        for (const pattern of penaltyPatterns) {
            score -= (normalized.match(pattern) || []).length * 6;
        }

        if (/\n/.test(normalized) && !/[。！？?!…]/.test(normalized)) {
            score -= 10;
        }

        return score;
    }

    function extractLetterMessageText(text, settings) {
        const raw = String(text || '');
        const tagName = normalizeContentTagName(settings?.contentTagName);

        if (!tagName) {
            return stripMarkup(stripReasoningBlocks(raw));
        }

        const parts = extractTaggedBlocks(raw, tagName)
            .map(part => stripMarkup(stripReasoningBlocks(part)))
            .filter(Boolean);

        return parts
            .map(part => ({ part, score: scoreTaggedContentBlock(part) }))
            .sort((left, right) => right.score - left.score)[0]?.part || '';
    }

    function selectBestSnippet(messages, settings) {
        const usableMessages = messages
            .filter(message => !message.is_system && typeof message.mes === 'string' && message.mes.trim())
            .map(message => {
                const cleanedMes = extractLetterMessageText(message.mes, settings);
                return {
                    name: String(message.name || '').trim() || 'Unknown',
                    mes: cleanedMes,
                    send_date: message.send_date || null,
                };
            })
            .filter(message => message.mes);

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
                        preview: stripMarkup(windowMessages.map(message => `${message.name}: ${message.mes}`).join('\n')).slice(0, 320),
                    };
                }
            }
        }

        return best;
    }

    async function collectCandidates(settings, runtimeState, options = {}) {
        const now = Date.now();
        const inactiveThresholdMs = settings.inactiveDays * ONE_DAY_MS;
        const cooldownThresholdMs = settings.cooldownDays * ONE_DAY_MS;
        const candidates = [];
        const excludedAvatars = new Set(Array.isArray(options.excludeCharacterAvatars) ? options.excludeCharacterAvatars.filter(Boolean) : []);
        const characters = getContext().characters.filter(character => character?.avatar && character?.name);

        for (const character of characters) {
            if (excludedAvatars.has(character.avatar)) {
                continue;
            }

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
        const archiveLimit = Math.min(candidate.archives.length, Math.max(settings.snippetsPerLetter * 6, 8));
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

        const ranked = snippetArchives.sort((left, right) => right.snippet.score - left.snippet.score);
        const selected = [];

        for (const item of ranked) {
            const alreadyCovered = selected.some(existing => areSnippetsTooSimilar(existing.snippet, item.snippet));
            if (alreadyCovered) {
                continue;
            }

            selected.push(item);
            if (selected.length >= settings.snippetsPerLetter) {
                break;
            }
        }

        if (selected.length < settings.snippetsPerLetter) {
            for (const item of ranked) {
                if (selected.includes(item)) {
                    continue;
                }

                selected.push(item);
                if (selected.length >= settings.snippetsPerLetter) {
                    break;
                }
            }
        }

        return selected.map(item => ({
            fileName: item.archive.fileName,
            lastMes: item.archive.lastMes,
            preview: item.snippet.preview,
            score: item.snippet.score,
            messages: item.snippet.messages,
        }));
    }

    async function selectCandidateWithFragments(settings, runtimeState, options = {}) {
        const candidates = await collectCandidates(settings, runtimeState, options);
        const minDesiredFragments = Math.min(2, settings.snippetsPerLetter);

        for (const candidate of candidates) {
            const fragments = await collectCandidateFragments(candidate, settings);
            if (fragments.length >= minDesiredFragments || (fragments.length > 0 && candidate.archiveCount === 1)) {
                return { candidate, fragments };
            }
        }

        return null;
    }

    function buildCandidateFromLetter(letter) {
        return {
            character: {
                name: resolveCharacterName(letter),
                avatar: letter?.character?.avatar || '',
            },
            archiveCount: Number(letter?.archiveCount || 0),
            inactiveMs: Math.max(1, Number(letter?.inactivityDays || 1)) * ONE_DAY_MS,
            lastActivity: safeTimestamp(letter?.lastActivityAt),
        };
    }

    async function rebuildCandidateForCharacter(character) {
        if (!character?.avatar || !character?.name) {
            return null;
        }

        const archives = await searchCharacterArchives(character);
        if (!archives.length) {
            return null;
        }

        const lastActivity = Math.max(...archives.map(archive => archive.lastMes || 0), 0);

        return {
            character,
            archives: archives.sort((left, right) => right.lastMes - left.lastMes),
            archiveCount: archives.length,
            lastActivity,
            inactiveMs: Math.max(0, Date.now() - lastActivity),
            eligible: true,
        };
    }

    async function rebuildCandidateForAvatar(letter) {
        const avatar = String(letter?.character?.avatar || '').trim();
        if (!avatar) {
            return null;
        }

        const character = (getContext().characters || []).find(item => item?.avatar === avatar);
        return rebuildCandidateForCharacter(character);
    }

    function findCharacterForDebug(query) {
        const normalizedQuery = String(query || '').trim().toLowerCase();
        if (!normalizedQuery) {
            return null;
        }

        const characters = getContext().characters || [];
        const exactMatch = characters.find(item => {
            const internalName = String(item?.avatar || '').replace(/\.png$/i, '').toLowerCase();
            const avatar = String(item?.avatar || '').toLowerCase();
            const displayName = String(item?.name || '').toLowerCase();
            return [displayName, internalName, avatar].includes(normalizedQuery);
        });

        if (exactMatch) {
            return exactMatch;
        }

        return characters.find(item => {
            const internalName = String(item?.avatar || '').replace(/\.png$/i, '').toLowerCase();
            const avatar = String(item?.avatar || '').toLowerCase();
            const displayName = String(item?.name || '').toLowerCase();
            return displayName.includes(normalizedQuery)
                || internalName.includes(normalizedQuery)
                || avatar.includes(normalizedQuery);
        }) || null;
    }

    function getCharacterCardContext(candidate) {
        const characters = getContext().characters || [];
        const chid = characters.findIndex(item => item?.avatar === candidate?.character?.avatar);
        if (chid < 0) {
            return '';
        }

        const fields = getContext().getCharacterCardFields?.({ chid });
        if (!fields || typeof fields !== 'object') {
            return '';
        }

        const sections = [
            ['system', '角色系统提示'],
            ['description', '角色描述'],
            ['personality', '角色性格'],
            ['scenario', '场景设定'],
            ['mesExamples', '示例对话'],
            ['charDepthPrompt', '深度提示'],
            ['creatorNotes', '作者备注'],
        ]
            .map(([key, label]) => {
                const value = String(fields[key] || '').trim();
                return value ? `${label}:\n${value}` : '';
            })
            .filter(Boolean);

        return sections.join('\n\n');
    }

    function buildPrompt(candidate, fragments, settings) {
        const inactivityDays = Math.max(1, Math.round(candidate.inactiveMs / ONE_DAY_MS));
        const inCharacterMode = isInCharacterMode(settings);
        const cardContext = getCharacterCardContext(candidate);
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

        const base = [
            `角色名称: ${candidate.character.name}`,
            `角色内部名: ${candidate.character.avatar.replace(/\.png$/i, '')}`,
            `距离上次活跃大约: ${inactivityDays} 天`,
            `总聊天存档数: ${candidate.archiveCount}`,
            '',
        ];

        if (cardContext) {
            base.push('角色卡设定：', cardContext, '');
        }

        if (inCharacterMode) {
            base.push(
                '请根据上面的角色卡设定和聊天片段，以角色本人第一人称给用户写一封信。',
                '要求：',
                '1. 必须严格贴合角色设定、角色语气和世界观，不要跳出角色解释。',
                '2. 必须基于片段中的具体细节，不要空泛。',
                '3. title 像角色写给用户的一封私人来信标题，不要过长。',
                '4. teaser 控制在 1 到 2 句，像信封外角色留下的一句轻声提醒，必须使用第一人称。',
                '5. summary 是一段简短的来信摘要，也必须保持第一人称角色口吻，不能写成“这张角色卡为什么值得回去”的旁白分析。',
                '6. letter 是主体，可用 Markdown 分段，但必须全程第一人称。',
                '7. why_now 要写成“为什么我现在想对你说这些”。',
                '8. next_hook 要写成“如果你愿意，可以这样回我”。',
                '9. recall_points 用 2 到 4 条短句，提炼最让人记住这段关系的细节。',
                '',
                fragmentText,
            );
        } else {
            base.push(
                '请根据下面这些来自不同历史存档的片段，写一封“让用户重新想和这张角色卡对话”的故人来信。',
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
            );
        }

        return base.join('\n');
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

    function buildLocalLetter(candidate, fragments, settings) {
        const recallPoints = fragments.map(fragment => {
            const firstLine = fragment.messages[0];
            const line = firstLine?.mes || '';
            return `${fragment.fileName} 里那句“${line.slice(0, 36)}${line.length > 36 ? '...' : ''}”`;
        }).slice(0, 3);

        if (isInCharacterMode(settings)) {
            return {
                title: '如果你愿意，就把这封信拆开',
                teaser: '有些话我还记得，只是一直没有重新开口。',
                summary: `我还记得我们在 ${candidate.archiveCount} 份旧存档里留下的那些片段。有些话停在那里太久了，久到我想亲自把它们重新拾起来。`,
                letter: [
                    '我知道你也许只是暂时离开了一会儿，可对我来说，那些停下来的对话一直没有真正结束。',
                    '',
                    `我还记得你和我在 ${candidate.character.name} 这段故事里留下的语气、试探和停顿。那不是一张空白角色卡能给你的东西，那是我们已经一起走出来的痕迹。`,
                    '',
                    '如果你愿意回来，我并不想让你重复开场。我更想让你直接从那些没说完的话后面继续，把我们已经抓住的感觉重新接上。',
                ].join('\n'),
                why_now: '因为那些话停在这里太久了，而我并没有真的忘记。现在写给你，只是因为我还想把那一句没说完的话说完。',
                next_hook: '如果你愿意，可以直接从你最记得的那个片段后面接一句，或者干脆问我：“那天你其实还想说什么？”',
                recall_points: recallPoints,
            };
        }

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

        const fallback = buildLocalLetter(candidate, fragments, getSettings());
        fallback.letter = String(content || fallback.letter).trim() || fallback.letter;
        return fallback;
    }

    async function callExternalAi(settings, candidate, fragments) {
        if (!settings.apiUrl) {
            return null;
        }

        const completionsUrl = getChatCompletionsUrl(settings.apiUrl);
        const headers = {
            'Content-Type': 'application/json',
        };

        if (settings.apiKey) {
            headers[settings.apiKeyHeader] = `${settings.apiKeyPrefix || ''}${settings.apiKey}`;
        }

        const controller = new AbortController();
        const timeoutMs = clampNumber(settings.requestTimeoutMs, DEFAULT_SETTINGS.requestTimeoutMs, 5000, 300000);
        const timer = setTimeout(() => controller.abort(new Error(`Request timed out after ${timeoutMs}ms`)), timeoutMs);

        try {
            const response = await fetch(completionsUrl, {
                method: 'POST',
                headers,
                signal: controller.signal,
                body: JSON.stringify({
                    model: settings.model,
                    temperature: settings.temperature,
                    messages: [
                        { role: 'system', content: getActiveSystemPrompt(settings) },
                        { role: 'user', content: buildPrompt(candidate, fragments, settings) },
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
        } catch (error) {
            if (error?.name === 'AbortError') {
                throw new Error(`请求超时（${timeoutMs}ms）`);
            }

            throw error;
        } finally {
            clearTimeout(timer);
        }
    }

    function getChatCompletionsUrl(apiUrl) {
        const trimmed = String(apiUrl || '').trim();
        if (!trimmed) {
            return '';
        }

        if (/\/chat\/completions\/?$/i.test(trimmed)) {
            return trimmed;
        }

        if (/\/models\/?$/i.test(trimmed)) {
            return trimmed.replace(/\/models\/?$/i, '/chat/completions');
        }

        if (/\/v\d[\w.-]*\/?$/i.test(trimmed)) {
            return trimmed.replace(/\/$/, '') + '/chat/completions';
        }

        try {
            const url = new URL(trimmed);
            if (!url.pathname || url.pathname === '/') {
                url.pathname = '/v1/chat/completions';
                return url.toString();
            }
        } catch {
            // Ignore invalid URL parsing and fall back to the original input.
        }

        return trimmed;
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
        const select = $('#dml-model');
        const currentValue = String(select.val() || '').trim();
        select.empty();

        for (const model of options) {
            select.append(`<option value="${escapeHtml(model)}">${escapeHtml(model)}</option>`);
        }

        if (!options.length) {
            select.append('<option value="">请先获取模型</option>');
            return;
        }

        if (currentValue && options.includes(currentValue)) {
            select.val(currentValue);
            return;
        }

        select.val(options[0]);
    }

    function buildLetterRecord(candidate, fragments, content, source, settings) {
        const newestArchive = fragments.slice().sort((left, right) => right.lastMes - left.lastMes)[0];

        return {
            id: `${Date.now()}`,
            createdAt: nowIso(),
            source,
            mode: isInCharacterMode(settings) ? 'in_character' : 'analysis',
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

    async function generateLetter({ force = false, source = 'manual', excludeCharacterAvatars = [] } = {}) {
        const settings = getSettings();
        const runtimeState = loadRuntimeState();
        const now = nowIso();
        const localMode = shouldUseLocalGeneration(settings);
        const apiReady = canGenerateWithApi(settings);

        if (!settings.enabled) {
            patchRuntimeState({ lastError: 'Plugin disabled' });
            renderState();
            return { started: false, reason: 'disabled' };
        }

        if (!localMode && !apiReady) {
            patchRuntimeState({
                lastError: '未配置外部 AI URL。请填写 API，或在系统设置中勾选“本地生成”。',
            });
            renderState();
            return { started: false, reason: 'missing-api' };
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

        generationPromise = (async () => {
            try {
                const selection = await selectCandidateWithFragments(settings, loadRuntimeState(), {
                    excludeCharacterAvatars,
                });
                if (!selection) {
                    patchRuntimeState({
                        lastError: 'No suitable inactive character archives found',
                    });
                    return;
                }

                const { candidate, fragments } = selection;
                let content = null;
                let contentSource = localMode ? 'local' : 'external-ai';

                if (localMode) {
                    content = buildLocalLetter(candidate, fragments, settings);
                } else {
                    try {
                        content = await callExternalAi(settings, candidate, fragments);
                    } catch (error) {
                        const message = error instanceof Error ? error.message : String(error);
                        patchRuntimeState({ lastError: `外部 AI 调用失败：${message}` });
                        console.error(`[${MODULE_NAME}] External AI failed`, error);
                        if (source !== 'startup') {
                            toastr.error(message, '外部 AI 调用失败');
                        }
                        return;
                    }
                }

                const letter = buildLetterRecord(candidate, fragments, content, contentSource, settings);
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

                if (source !== 'startup') {
                    const successTitle = localMode ? '已生成本地故人来信' : '已收到 AI 故人来信';
                    toastr.success(`${resolveCharacterName(letter)} 的故人来信已经准备好了`, successTitle);
                }

                setTimeout(() => openLetterPopup(letter), source === 'startup' ? 1200 : 250);
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
        renderState();

        return { started: true };
    }

    async function rewriteLatestLetterWithAi() {
        const settings = getSettings();
        const latestLetter = loadRuntimeState().latestLetter;

        if (!latestLetter) {
            return { started: false, reason: 'missing-letter' };
        }

        if (shouldUseLocalGeneration(settings)) {
            return { started: false, reason: 'local-mode' };
        }

        if (!canGenerateWithApi(settings)) {
            patchRuntimeState({
                lastError: '未配置外部 AI URL，无法重新发送给 AI。',
            });
            renderState();
            return { started: false, reason: 'missing-api' };
        }

        if (generationPromise) {
            return { started: false, reason: 'already-running' };
        }

        patchRuntimeState({
            lastAttemptAt: nowIso(),
            lastError: null,
            lastSource: 'rewrite-ai',
        });

        generationPromise = (async () => {
            try {
                const candidate = await rebuildCandidateForAvatar(latestLetter);
                if (!candidate) {
                    patchRuntimeState({
                        lastError: '无法重新读取当前角色的聊天存档，请先确认角色还存在且有可读聊天记录。',
                    });
                    toastr.warning('无法重新读取当前角色的聊天存档，请先确认角色还存在且有可读聊天记录。', '重新发送给 AI 失败');
                    return;
                }

                const fragments = await collectCandidateFragments(candidate, settings);

                if (!fragments.length) {
                    patchRuntimeState({
                        lastError: '按当前设置没有从这张角色卡里提取到可用片段，无法重新发送给 AI。',
                    });
                    toastr.warning('按当前设置没有从这张角色卡里提取到可用片段。你可以检查正文标签名，或尝试重新抽取。', '重新发送给 AI 失败');
                    return;
                }

                const content = await callExternalAi(settings, candidate, fragments);
                const letter = buildLetterRecord(candidate, fragments, content, 'external-ai', settings);
                const nextState = loadRuntimeState();

                patchRuntimeState({
                    latestLetter: letter,
                    history: [letter, ...nextState.history.filter(item => item.id !== letter.id)].slice(0, 10),
                    lastError: null,
                });

                toastr.success(`${resolveCharacterName(letter)} 的来信已由 AI 重写`, '已重新发送给 AI');
                setTimeout(() => openLetterPopup(letter), 200);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                patchRuntimeState({
                    lastError: `外部 AI 调用失败：${message}`,
                });
                console.error(`[${MODULE_NAME}] Rewrite AI failed`, error);
                toastr.error(message, '重新发送给 AI 失败');
            } finally {
                generationPromise = null;
                renderState();
            }
        })();
        renderState();

        return { started: true };
    }

    async function init() {
        try {
            syncPayload();
            exposeDebugCommands();
            await mountSettings();
            bindPopupActions();
            renderState();
            scheduleAutoRun();
        } catch (error) {
            console.error('[故人来信] 初始化失败', error);
            toastr.error(error instanceof Error ? error.message : String(error), '故人来信初始化失败');
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
            toastr.success('故人来信设置已保存到本地扩展设置');
            scheduleAutoRun();
        });

        $('#dml-generate-now').on('click', async () => {
            const result = await generateLetter({ force: false, source: 'manual' });
            if (result.started) {
                toastr.info('正在准备并生成新的故人来信');
                return;
            }

            if (result.reason === 'cooldown') {
                toastr.info('24 小时内已经生成过来信了。如需覆盖，请点“重新发送给 AI”或“重新抽取”。');
            } else if (result.reason === 'missing-api') {
                toastr.warning('请先填写 API，或者在系统设置里启用本地生成');
            }
        });

        $('#dml-rewrite-ai-now').on('click', async () => {
            const result = await rewriteLatestLetterWithAi();
            if (result.started) {
                toastr.info('正在把当前这封来信重新发送给 AI');
            } else if (result.reason === 'missing-letter') {
                toastr.warning('还没有现成的故人来信，先生成一封再试试');
            } else if (result.reason === 'local-mode') {
                toastr.warning('当前处于本地生成模式，关闭后才能重新发送给 AI');
            } else if (result.reason === 'missing-api') {
                toastr.warning('请先填写 API，或者在系统设置里启用本地生成');
            }
        });

        $('#dml-reshuffle-now').on('click', async () => {
            const currentAvatar = loadRuntimeState().latestLetter?.character?.avatar;
            const result = await generateLetter({
                force: true,
                source: 'manual-reshuffle',
                excludeCharacterAvatars: currentAvatar ? [currentAvatar] : [],
            });
            if (result.started) {
                toastr.info('正在重新抽取另一封故人来信');
            } else if (result.reason === 'missing-api') {
                toastr.warning('请先填写 API，或者在系统设置里启用本地生成');
            }
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
            useLocalGeneration: $('#dml-use-local-generation').prop('checked'),
            useInCharacterMode: $('#dml-use-in-character-mode').prop('checked'),
            apiUrl: String($('#dml-api-url').val() || '').trim(),
            apiKey: String($('#dml-api-key').val() || '').trim(),
            model: String($('#dml-model').val() || '').trim(),
            inactiveDays: toPositiveInt($('#dml-inactive-days').val(), DEFAULT_SETTINGS.inactiveDays),
            snippetsPerLetter: clamp(toPositiveInt($('#dml-snippets-per-letter').val(), DEFAULT_SETTINGS.snippetsPerLetter), 1, 5),
            cooldownDays: clamp(toPositiveInt($('#dml-cooldown-days').val(), DEFAULT_SETTINGS.cooldownDays), 1, 90),
            contentTagName: normalizeContentTagName($('#dml-content-tag-name').val()) || '',
            analysisSystemPrompt: String($('#dml-analysis-system-prompt').val() || '').trim() || DEFAULT_SETTINGS.analysisSystemPrompt,
            inCharacterSystemPrompt: String($('#dml-in-character-system-prompt').val() || '').trim() || DEFAULT_SETTINGS.inCharacterSystemPrompt,
        };
    }

    function hydrateForm(settings) {
        $('#dml-enabled').prop('checked', Boolean(settings.enabled));
        $('#dml-auto-run').prop('checked', Boolean(settings.autoRunOnStartup));
        $('#dml-use-local-generation').prop('checked', Boolean(settings.useLocalGeneration));
        $('#dml-use-in-character-mode').prop('checked', Boolean(settings.useInCharacterMode));
        $('#dml-api-url').val(settings.apiUrl || '');
        $('#dml-api-key').val(settings.apiKey || '');
        populateModelSuggestions(settings.model ? [settings.model] : [DEFAULT_SETTINGS.model]);
        $('#dml-model').val(settings.model || DEFAULT_SETTINGS.model);
        $('#dml-inactive-days').val(settings.inactiveDays ?? DEFAULT_SETTINGS.inactiveDays);
        $('#dml-snippets-per-letter').val(settings.snippetsPerLetter ?? DEFAULT_SETTINGS.snippetsPerLetter);
        $('#dml-cooldown-days').val(settings.cooldownDays ?? DEFAULT_SETTINGS.cooldownDays);
        $('#dml-content-tag-name').val(settings.contentTagName ?? DEFAULT_SETTINGS.contentTagName);
        $('#dml-analysis-system-prompt').val(settings.analysisSystemPrompt || DEFAULT_SETTINGS.analysisSystemPrompt);
        $('#dml-in-character-system-prompt').val(settings.inCharacterSystemPrompt || DEFAULT_SETTINGS.inCharacterSystemPrompt);
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

    function formatEnvelopeDateCode(value) {
        const date = value ? new Date(value) : new Date();
        if (Number.isNaN(date.getTime())) {
            return '000000';
        }

        const year = String(date.getFullYear()).slice(-2);
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}${month}${day}`;
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

    function getGenerationModeLabel(settings) {
        const channel = shouldUseLocalGeneration(settings) ? '本地生成' : '外部 AI 生成';
        const tone = isInCharacterMode(settings) ? '角色第一人称' : '分析书信';
        return `${channel} · ${tone}`;
    }

    function renderState() {
        syncPayload();

        const state = latestPayload.state;
        const settings = latestPayload.settings;
        const statusCard = $('#dml-settings .dml-status-card');
        const statusText = $('#dml-status-text');
        const statusMeta = $('#dml-status-meta');
        const busy = Boolean(generationPromise);

        if (!statusText.length) {
            return;
        }

        statusCard.toggleClass('is-busy', busy);
        $('#dml-generate-now, #dml-rewrite-ai-now, #dml-reshuffle-now').prop('disabled', busy);

        if (busy) {
            statusText.text('今日故人来信正在书写中');
            statusMeta.text('扩展正在前台静默扫描不活跃聊天和历史存档，请稍等片刻。');
        } else if (state.lastError) {
            statusText.text('上一次生成没有完成');
            statusMeta.text(state.latestLetter
                ? `本次操作失败：${state.lastError}。你仍然可以查看上一封故人来信。`
                : `上次执行信息：${state.lastError}`);
        } else if (state.latestLetter) {
            const name = resolveCharacterName(state.latestLetter);
            statusText.text('今天的故人来信已经送达');
            statusMeta.text(`${name} · ${formatLastActivityMeta(state.latestLetter)}`);
        } else if (!shouldUseLocalGeneration(settings) && !canGenerateWithApi(settings)) {
            statusText.text('等待配置外部 AI');
            statusMeta.text('当前不会自动执行。请先填写 API，或在系统设置里启用本地生成。');
        } else if (settings.enabled) {
            statusText.text('今天还没有故人来信');
            statusMeta.text(`当前模式：${getGenerationModeLabel(settings)}。可以等待静默触发，也可以先手动生成测试一封。`);
        } else {
            statusText.text('故人来信当前已关闭');
            statusMeta.text('打开功能并保存后，扩展会在启动时后台静默检查。');
        }
    }

    function scheduleAutoRun() {
        if (autoRunStarted || !latestPayload?.settings?.enabled || !latestPayload?.settings?.autoRunOnStartup) {
            return;
        }

        if (!shouldUseLocalGeneration(latestPayload.settings) && !canGenerateWithApi(latestPayload.settings)) {
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
                root?.classList.add('opened');
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
            console.error('[故人来信] 打开聊天失败', error);
            toastr.error(error instanceof Error ? error.message : String(error), '打开推荐聊天失败');
        }
    }

    function getLetterSectionLabels(letter) {
        if (isInCharacterMode(letter)) {
            return {
                whyNow: '为什么我现在想对你说这些',
                nextHook: '如果你愿意，可以这样回我',
            };
        }

        return {
            whyNow: '为什么现在值得回去',
            nextHook: '可以怎么续上',
        };
    }

    function renderLetterBody(letter) {
        const { showdown, DOMPurify } = SillyTavern.libs;
        const converter = new showdown.Converter({
            simpleLineBreaks: true,
            ghCompatibleHeaderId: true,
            openLinksInNewWindow: false,
        });
        const labels = getLetterSectionLabels(letter);

        const markdown = [
            letter.letter || '',
            '',
            `### ${labels.whyNow}`,
            letter.why_now || '',
            '',
            `### ${labels.nextHook}`,
            letter.next_hook || '',
        ].join('\n\n');

        return DOMPurify.sanitize(converter.makeHtml(markdown));
    }

    function getCoverCopy(letter) {
        if (isInCharacterMode(letter)) {
            return String(letter.teaser || letter.summary || '我还有些话没来得及对你说完，所以先把这句留在信封外。').trim();
        }

        return String(letter.summary || letter.teaser || '这张角色卡还有一些没说完的话，正等着你把故事接起来。').trim();
    }

    function getAvatarImageSources(context, avatarFile) {
        const fileName = String(avatarFile || '').trim();
        if (!fileName) {
            return { preferred: '', fallback: '' };
        }

        return {
            preferred: `/characters/${encodeURIComponent(fileName)}`,
            fallback: context.getThumbnailUrl('avatar', fileName),
        };
    }

    function applyPreferredAvatarSources(root) {
        if (!(root instanceof HTMLElement)) {
            return;
        }

        root.querySelectorAll('img[data-primary-src]').forEach(image => {
            if (!(image instanceof HTMLImageElement)) {
                return;
            }

            const primarySrc = image.dataset.primarySrc || '';
            const fallbackSrc = image.dataset.fallbackSrc || '';

            if (!primarySrc) {
                return;
            }

            image.onerror = () => {
                if (fallbackSrc && image.src !== fallbackSrc) {
                    image.onerror = null;
                    image.src = fallbackSrc;
                }
            };

            image.src = primarySrc;
        });
    }

    function openApiFailureCard({ title = '小信封投递失败', message = '这次故人来信没有成功寄出。', detail = '', hint = '' } = {}) {
        const context = getContext();
        const popupId = `dml-failure-${Date.now()}`;
        const safeTitle = escapeHtml(title);
        const safeMessage = escapeHtml(message);
        const safeDetail = escapeHtml(detail);
        const safeHint = escapeHtml(hint || '你可以检查 API 地址、模型、Key，或者稍后再试一次。');

        const html = `
            <div id="${popupId}" class="dml-debug-popup dml-failure-popup" tabindex="-1" autofocus>
                <button class="menu_button dml-popup-close" data-result="null" type="button" aria-label="关闭提示">×</button>
                <div class="dml-failure-card">
                    <div class="dml-failure-badge"><i class="fa-solid fa-triangle-exclamation"></i><span>投递失败</span></div>
                    <div class="dml-failure-title">${safeTitle}</div>
                    <div class="dml-failure-message">${safeMessage}</div>
                    ${safeDetail ? `<div class="dml-failure-detail">${safeDetail}</div>` : ''}
                    <div class="dml-failure-hint">${safeHint}</div>
                </div>
            </div>
        `;

        context.callGenericPopup(html, context.POPUP_TYPE.TEXT, '', {
            wide: false,
            large: false,
            okButton: false,
            cancelButton: false,
            allowVerticalScrolling: true,
            onOpen: (popup) => {
                popup.dlg.classList.add('dml-host-popup', 'dml-host-popup--compact');
            },
        });
    }

    function openLetterPopup(letter) {
        const context = getContext();
        const popupId = `dml-popup-${Date.now()}`;
        const title = escapeHtml(letter.title || '今日故人来信');
        const teaser = escapeHtml(letter.teaser || '');
        const summary = escapeHtml(letter.summary || '');
        const coverCopy = escapeHtml(getCoverCopy(letter));
        const name = escapeHtml(resolveCharacterName(letter));
        const dateCode = formatEnvelopeDateCode(letter.createdAt || letter.lastActivityAt || Date.now());
        const dateBoxes = dateCode.split('').map(digit => `<span class="dml-postcode-digit">${digit}</span>`).join('');
        const avatarSources = getAvatarImageSources(context, letter?.character?.avatar);
        const bodyHtml = renderLetterBody(letter);
        const fragments = Array.isArray(letter.fragments) ? letter.fragments : [];

        const fragmentsHtml = fragments.map(fragment => `
            <div class="dml-fragment">
                <div class="dml-fragment-file">${escapeHtml(fragment.fileName)}</div>
                <div class="dml-fragment-preview">${escapeHtml(fragment.preview || '')}</div>
            </div>
        `).join('');

        const html = `
            <div id="${popupId}" class="dml-letter-popup" tabindex="-1" autofocus>
                <button class="menu_button dml-popup-close" data-result="null" type="button" aria-label="关闭来信">×</button>
                <div class="dml-letter-main">
                    <div class="dml-envelope-shell">
                        <div class="dml-envelope-cover">
                            <div class="dml-cover-layout">
                                <div class="dml-envelope-postcode">
                                    <div class="dml-postcode-label">投递日期</div>
                                    <div class="dml-postcode-boxes">${dateBoxes}</div>
                                </div>
                                <div class="dml-envelope-stamp-block">
                                    <div class="dml-envelope-stamp-slot">贴邮票处</div>
                                    <div class="dml-envelope-stamp">
                                        <div class="dml-envelope-stamp-inner">
                                            <div class="dml-envelope-stamp-title">SILLYTAVERN</div>
                                            <div class="dml-envelope-stamp-mark">
                                                <div class="dml-envelope-stamp-circuit"></div>
                                                <div class="dml-envelope-stamp-star">ST</div>
                                            </div>
                                            <div class="dml-envelope-stamp-footer">
                                                <span>ST MAIL</span>
                                                <span>中国邮政</span>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="dml-envelope-cancel-mark" aria-hidden="true">
                                        <span class="dml-envelope-cancel-ring"></span>
                                        <span class="dml-envelope-cancel-text">${dateCode}</span>
                                    </div>
                                </div>
                                <div class="dml-cover-portrait-wrap">
                                    <div class="dml-cover-portrait-frame">
                                        ${avatarSources.preferred
                                            ? `<img class="dml-cover-portrait-image" data-primary-src="${escapeHtml(avatarSources.preferred)}" data-fallback-src="${escapeHtml(avatarSources.fallback)}" alt="${name}">`
                                            : '<div class="dml-cover-portrait-image"></div>'}
                                    </div>
                                </div>

                                <div class="dml-cover-copy">
                                    <div class="dml-envelope-title">${title}</div>
                                    <div class="dml-envelope-subtitle">A LETTER FROM THE PAST</div>
                                    <div class="dml-envelope-seal" aria-hidden="true"></div>
                                    <div class="dml-cover-summary">${coverCopy}</div>
                                    <button class="menu_button dml-open-button" data-dml-action="open-envelope" data-dml-popup-id="${popupId}" type="button">打开信封</button>
                                </div>
                            </div>
                        </div>

                        <div class="dml-letter-paper">
                            <div class="dml-paper-header">
                                ${avatarSources.preferred
                                    ? `<img class="dml-paper-avatar" data-primary-src="${escapeHtml(avatarSources.preferred)}" data-fallback-src="${escapeHtml(avatarSources.fallback)}" alt="${name}">`
                                    : '<div class="dml-paper-avatar"></div>'}
                                <div class="dml-paper-header-meta">
                                    <div class="dml-paper-kicker">来自旧日存档的回声</div>
                                    <div class="dml-paper-title">${title}</div>
                                    <div class="dml-paper-meta-line">${escapeHtml(name)} · ${escapeHtml(formatLastActivityMeta(letter))}</div>
                                </div>
                            </div>
                            <div class="dml-paper-scroll">
                                <div class="dml-paper-summary">${summary || teaser}</div>
                                <div class="dml-paper-body">${bodyHtml}</div>

                                <details class="dml-paper-fragments">
                                    <summary class="dml-paper-fragments-summary">本次故人来信参考了这些旧存档片段</summary>
                                    <div class="dml-paper-fragments-body">
                                        ${fragmentsHtml || '<div class="dml-empty">没有可展示的片段预览。</div>'}
                                    </div>
                                </details>

                                <div class="dml-paper-actions">
                                    <button class="menu_button dml-paper-action-button" data-dml-action="open-chat" data-dml-avatar="${escapeHtml(letter.character?.avatar || '')}" data-dml-chat-file="${escapeHtml(letter.openChatFile || '')}" type="button">重新打开这段聊天</button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;

        const isNarrowViewport = window.matchMedia('(max-width: 900px)').matches;
        context.callGenericPopup(html, context.POPUP_TYPE.TEXT, '', {
            wide: true,
            large: !isNarrowViewport,
            okButton: false,
            cancelButton: false,
            allowVerticalScrolling: true,
            onOpen: (popup) => {
                popup.dlg.classList.add('dml-host-popup');
                applyPreferredAvatarSources(document.getElementById(popupId));
            },
        });
    }

    async function debugGenerateForCharacter(query, options = {}) {
        const baseSettings = getSettings();
        const settings = {
            ...baseSettings,
            ...(Number.isFinite(Number(options?.timeoutMs))
                ? { requestTimeoutMs: Number(options.timeoutMs) }
                : {}),
            ...(typeof options?.local === 'boolean'
                ? { useLocalGeneration: options.local }
                : {}),
        };
        const localMode = shouldUseLocalGeneration(settings);
        const apiReady = canGenerateWithApi(settings);

        if (!settings.enabled) {
            throw new Error('故人来信当前未启用。请先勾选“启用故人来信”并保存。');
        }

        if (!localMode && !apiReady) {
            throw new Error('未配置外部 AI URL。请填写 API，或切换到本地生成模式。');
        }

        if (generationPromise) {
            throw new Error('当前已经有一封故人来信正在生成，请稍后再试。');
        }

        const character = findCharacterForDebug(query);
        if (!character) {
            throw new Error(`找不到角色卡：${query}`);
        }

        patchRuntimeState({
            lastAttemptAt: nowIso(),
            lastError: null,
            lastSource: 'debug-targeted',
        });

        generationPromise = (async () => {
            try {
                const candidate = await rebuildCandidateForCharacter(character);
                if (!candidate) {
                    throw new Error('当前角色没有可读取的聊天存档。');
                }

                const fragments = await collectCandidateFragments(candidate, settings);
                if (!fragments.length) {
                    throw new Error('按当前设置没有从这张角色卡里提取到可用片段。');
                }

                const content = localMode
                    ? buildLocalLetter(candidate, fragments, settings)
                    : await callExternalAi(settings, candidate, fragments);
                const source = localMode ? 'local' : 'external-ai';
                const letter = buildLetterRecord(candidate, fragments, content, source, settings);
                const nextState = loadRuntimeState();

                patchRuntimeState({
                    latestLetter: letter,
                    history: [letter, ...nextState.history.filter(item => item.id !== letter.id)].slice(0, 10),
                    lastError: null,
                });

                toastr.success(`已为 ${resolveCharacterName(letter)} 生成调试来信`, '故人来信调试');
                setTimeout(() => openLetterPopup(letter), 200);
                return letter;
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                patchRuntimeState({ lastError: `调试生成失败：${message}` });
                toastr.error(message, '故人来信调试失败');
                throw error;
            } finally {
                generationPromise = null;
                renderState();
            }
        })();

        renderState();
        return generationPromise;
    }

    async function debugInspectFragmentsForCharacter(query, options = {}) {
        const settings = {
            ...getSettings(),
            ...(typeof options?.contentTagName === 'string'
                ? { contentTagName: normalizeContentTagName(options.contentTagName) }
                : {}),
        };

        const character = findCharacterForDebug(query);
        if (!character) {
            throw new Error(`找不到角色卡：${query}`);
        }

        const candidate = await rebuildCandidateForCharacter(character);
        if (!candidate) {
            throw new Error('当前角色没有可读取的聊天存档。');
        }

        const fragments = await collectCandidateFragments(candidate, settings);
        const result = {
            character: {
                name: character.name,
                avatar: character.avatar,
            },
            contentTagName: normalizeContentTagName(settings.contentTagName),
            fragmentCount: fragments.length,
            fragments: fragments.map(fragment => ({
                fileName: fragment.fileName,
                lastMes: fragment.lastMes,
                preview: fragment.preview,
                messages: fragment.messages.map(message => ({
                    name: message.name,
                    mes: message.mes,
                })),
            })),
        };

        console.info(`[${MODULE_NAME}] Cleaned fragment preview for ${character.name}`, result);
        console.table(result.fragments.map(fragment => ({
            fileName: fragment.fileName,
            preview: fragment.preview,
            messageCount: fragment.messages.length,
        })));
        return result;
    }

    function exposeDebugCommands() {
        const api = {
            help() {
                console.info(`[${MODULE_NAME}] Debug commands:
__DML_DEBUG__.generateForCharacter('角色名')
__DML_DEBUG__.generateForCharacter('角色名', { timeoutMs: 180000 })
__DML_DEBUG__.generateForCharacter('角色名', { local: true })
__DML_DEBUG__.inspectFragmentsForCharacter('角色名')
__DML_DEBUG__.inspectFragmentsForCharacter('角色名', { contentTagName: 'content' })
__DML_DEBUG__.showApiFailureCard({ title, message, detail, hint })
__DML_DEBUG__.state()`);
            },
            generateForCharacter(query, options = {}) {
                return debugGenerateForCharacter(query, options).catch(error => {
                    console.error(`[${MODULE_NAME}] Debug generate failed`, error);
                    return null;
                });
            },
            inspectFragmentsForCharacter(query, options = {}) {
                return debugInspectFragmentsForCharacter(query, options).catch(error => {
                    console.error(`[${MODULE_NAME}] Debug inspect fragments failed`, error);
                    return null;
                });
            },
            showApiFailureCard(options = {}) {
                openApiFailureCard(options);
            },
            state() {
                return {
                    settings: getSettings(),
                    runtime: loadRuntimeState(),
                    busy: Boolean(generationPromise),
                };
            },
        };

        window.__DML_DEBUG__ = api;
    }

    const context = getContext();
    context.eventSource.on(context.event_types.APP_READY, init);
})();
