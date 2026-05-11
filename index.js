// ============================================================
// World Info Suite - Combined Extension for SillyTavern
// Features:
//   1. Triggered Entry Viewer - View triggered World Info per message
//   2. Character Lorebook Quick Access - Quick jump to character's lorebooks
//   3. Bulk Entry Editor - Batch edit World Info entries
// ============================================================

import {
  eventSource,
  event_types,
  chat,
  chat_metadata,
  characters,
  this_chid,
  saveSettingsDebounced,
} from '../../../../script.js';

import {
  extension_settings,
  renderExtensionTemplateAsync,
} from '../../../extensions.js';

import {
  callGenericPopup,
  POPUP_TYPE,
  Popup,
  POPUP_RESULT,
} from '../../../popup.js';

import {
  METADATA_KEY,
  selected_world_info,
  world_info,
  world_names,
  openWorldInfoEditor,
  createWorldInfoEntry,
  deleteWorldInfoEntry,
  loadWorldInfo,
  saveWorldInfo,
  newWorldInfoEntryDefinition,
  moveWorldInfoEntry,
} from '../../../world-info.js';

import { getCharaFilename, delay } from '../../../utils.js';
import { addLocaleData, getCurrentLocale, t } from '../../../i18n.js';

// ===== Extension Info =====
const url = new URL(import.meta.url);
const extensionName = url.pathname.substring(url.pathname.lastIndexOf('extensions/') + 11, url.pathname.lastIndexOf('/'));
const extensionFolderPath = `scripts/extensions/${extensionName}`;

// ===== Default Settings =====
const defaultSettings = {
  enableTriggeredViewer: true,
  enableCharLorebook: true,
  enableBulkEditor: true,
  enableWorldbookManager: true,
  viewerCacheLimit: 10, // Maximum number of messages to keep World Info viewer data
  viewerIcon: 'fa-globe', // Icon for the viewer button
  showGlobalLorebookMobile: true, // Show global lorebooks on mobile
  showGlobalLorebookDesktop: true, // Show global lorebooks on desktop
};

const WORLDBOOK_SORT_MODE = {
  CREATED_DESC: 'created_desc',
  CREATED_ASC: 'created_asc',
  NAME_ASC: 'name_asc',
  NAME_DESC: 'name_desc',
  CUSTOM: 'custom',
};

function createWorldbookManagerDefaults() {
  return {
    baseSort: WORLDBOOK_SORT_MODE.CREATED_DESC,
    customOrder: [],
    priorityKeywords: [],
    prioritizeCharacterBound: false,
    hiddenKeywords: [],
    worldTags: {},
    activeTagFilter: '',
  };
}

// ===== i18n System =====
let localeData = {};

async function loadLocaleData() {
  const locale = getCurrentLocale();
  const localeToLoad = locale.startsWith('zh') ? 'zh-tw' : 'en';

  try {
    const response = await fetch(`${extensionFolderPath}/locales/${localeToLoad}.json`);
    if (response.ok) {
      localeData = await response.json();
      // Add to SillyTavern's i18n system
      addLocaleData(locale, localeData);
      console.log(`[${extensionName}] Loaded locale: ${localeToLoad}`);
    }
  } catch (error) {
    console.warn(`[${extensionName}] Failed to load locale ${localeToLoad}, falling back to en`);
    try {
      const fallbackResponse = await fetch(`${extensionFolderPath}/locales/en.json`);
      if (fallbackResponse.ok) {
        localeData = await fallbackResponse.json();
      }
    } catch (e) {
      console.error(`[${extensionName}] Failed to load fallback locale`);
    }
  }
}

function i18n(key, ...args) {
  let text = localeData[key] || key;
  // Replace {0}, {1}, etc. with arguments
  args.forEach((arg, index) => {
    text = text.replace(new RegExp(`\\{${index}\\}`, 'g'), arg);
  });
  return text;
}

// ===== Position & Status Info (with i18n) =====
function getPositionInfo() {
  return {
    0: { name: i18n('positionBeforeCharDef'), emoji: '📙' },
    1: { name: i18n('positionAfterCharDef'), emoji: '📙' },
    2: { name: i18n('positionBeforeAN'), emoji: '📝' },
    3: { name: i18n('positionAfterAN'), emoji: '📝' },
    4: { name: i18n('positionAtDepth'), emoji: '💉' },
    5: { name: i18n('positionBeforeExamples'), emoji: '📄' },
    6: { name: i18n('positionAfterExamples'), emoji: '📄' },
    7: { name: i18n('positionOutlet'), emoji: '➡️' },
  };
}

const POSITION_SORT_ORDER = {
  0: 0, 1: 1, 5: 2, 6: 3, 2: 4, 3: 5, 4: 6, 7: 7,
};

function getSelectiveLogicInfo() {
  return {
    0: i18n('selectiveLogicAndAny'),
    1: i18n('selectiveLogicNotAll'),
    2: i18n('selectiveLogicNotAny'),
    3: i18n('selectiveLogicAndAll'),
  };
}

const WI_SOURCE_KEYS = {
  GLOBAL: 'global',
  CHARACTER_PRIMARY: 'characterPrimary',
  CHARACTER_ADDITIONAL: 'characterAdditional',
  CHAT: 'chat',
};

function getWiSourceDisplay() {
  return {
    [WI_SOURCE_KEYS.GLOBAL]: i18n('sourceGlobal'),
    [WI_SOURCE_KEYS.CHARACTER_PRIMARY]: i18n('sourceCharacterPrimary'),
    [WI_SOURCE_KEYS.CHARACTER_ADDITIONAL]: i18n('sourceCharacterAdditional'),
    [WI_SOURCE_KEYS.CHAT]: i18n('sourceChat'),
  };
}

const ENTRY_SOURCE_TYPE = {
  ASSISTANT: 3,
  USER: 2,
  SYSTEM: 1,
};

// ===== Helper Functions =====
function getRoleString(roleValue) {
  if (roleValue === 0) return 'system';
  if (roleValue === 1) return 'user';
  if (roleValue === 2) return 'assistant';
  if (typeof roleValue === 'string') {
    const lowerRole = roleValue.toLowerCase().trim();
    if (lowerRole === 'ai') return 'assistant';
    return lowerRole;
  }
  return 'assistant';
}

function roleDisplayName(role) {
  if (role === 'assistant') return i18n('roleAssistant');
  if (role === 'user') return i18n('roleUser');
  if (role === 'system') return i18n('roleSystem');
  return i18n('roleAssistant');
}

function getEntrySourceType(entry) {
  const role = getRoleString(entry.role);
  if (role === 'assistant') return ENTRY_SOURCE_TYPE.ASSISTANT;
  if (role === 'user') return ENTRY_SOURCE_TYPE.USER;
  if (role === 'system') return ENTRY_SOURCE_TYPE.SYSTEM;
  return ENTRY_SOURCE_TYPE.ASSISTANT;
}

function getWISourceKey(entry) {
  const worldName = entry.world;
  const chatLoreName = chat_metadata?.[METADATA_KEY];

  if (chatLoreName && worldName === chatLoreName) {
    return WI_SOURCE_KEYS.CHAT;
  }

  const character = characters?.[this_chid];
  if (character) {
    const primaryLorebook = character?.data?.extensions?.world;
    if (primaryLorebook && worldName === primaryLorebook) {
      return WI_SOURCE_KEYS.CHARACTER_PRIMARY;
    }

    const fileName = getCharaFilename?.(this_chid);
    const extraCharLore = world_info?.charLore?.find?.((e) => e.name === fileName);
    if (extraCharLore && Array.isArray(extraCharLore.extraBooks) && extraCharLore.extraBooks.includes(worldName)) {
      return WI_SOURCE_KEYS.CHARACTER_ADDITIONAL;
    }
  }

  if (Array.isArray(selected_world_info) && selected_world_info.includes(worldName)) {
    return WI_SOURCE_KEYS.GLOBAL;
  }

  return null;
}

function getSourceDisplayName(sourceKey) {
  if (!sourceKey) return '';
  return getWiSourceDisplay()[sourceKey] || '';
}

function getEntryStatus(entry) {
  if (entry.constant === true) return { emoji: '🔵', name: i18n('statusConstant') };
  if (entry.vectorized === true) return { emoji: '🔗', name: i18n('statusVectorized') };
  return { emoji: '🟢', name: i18n('statusKeyword') };
}

function formatRoleDepthTag(entry) {
  const roleString = getRoleString(entry.role);
  const depth = entry.depth ?? null;
  if (depth == null) return '';
  return `${roleDisplayName(roleString)} ${i18n('depthLabel')} ${depth}`;
}

const worldOrderCache = new Map();

function getWorldOrderByName(worldName) {
  if (!worldName) return Number.MAX_SAFE_INTEGER;
  if (worldOrderCache.has(worldName)) {
    return worldOrderCache.get(worldName);
  }

  let order = Number.MAX_SAFE_INTEGER;

  try {
    const candidates = [
      world_info?.worlds,
      world_info?.allWorlds,
      world_info?.files,
      world_info?.data,
      world_info?.all_worlds,
    ].filter(Boolean);

    for (const list of candidates) {
      if (Array.isArray(list)) {
        const found = list.find((w) => (w?.name || w?.title) === worldName);
        if (found && typeof found.order === 'number') {
          order = found.order;
          break;
        }
      } else if (typeof list === 'object') {
        const w = list[worldName];
        if (w && typeof w.order === 'number') {
          order = w.order;
          break;
        }
      }
    }
  } catch (err) { /* ignore */ }

  worldOrderCache.set(worldName, order);
  return order;
}

function getPositionSortIndex(position) {
  if (position in POSITION_SORT_ORDER) return POSITION_SORT_ORDER[position];
  return 999;
}

function compareDepthEntries(a, b) {
  const depthA = a.depth ?? -Infinity;
  const depthB = b.depth ?? -Infinity;
  if (depthA !== depthB) return depthB - depthA;

  const stA = a.sourceType ?? ENTRY_SOURCE_TYPE.ASSISTANT;
  const stB = b.sourceType ?? ENTRY_SOURCE_TYPE.ASSISTANT;
  if (stA !== stB) return stB - stA;

  const orderA = (typeof a.worldOrder === 'number') ? a.worldOrder : Number.MAX_SAFE_INTEGER;
  const orderB = (typeof b.worldOrder === 'number') ? b.worldOrder : Number.MAX_SAFE_INTEGER;
  if (orderA !== orderB) return orderA - orderB;

  return String(a.entryName || '').localeCompare(String(b.entryName || ''));
}

function compareOrderEntries(a, b) {
  const oa = (typeof a.worldOrder === 'number') ? a.worldOrder : Number.MAX_SAFE_INTEGER;
  const ob = (typeof b.worldOrder === 'number') ? b.worldOrder : Number.MAX_SAFE_INTEGER;
  if (oa !== ob) return oa - ob;

  const wn = String(a.worldName || '').localeCompare(String(b.worldName || ''));
  if (wn !== 0) return wn;
  return String(a.entryName || '').localeCompare(String(b.entryName || ''));
}

function normalizeKeywordToken(value) {
  return String(value ?? '').trim().toLocaleLowerCase();
}

function normalizeKeywordArray(value) {
  const arr = Array.isArray(value)
    ? value
    : (typeof value === 'string' ? value.split(/[,\n]/) : []);

  const seen = new Set();
  const result = [];

  for (const raw of arr) {
    const cleaned = String(raw ?? '').trim();
    if (!cleaned) continue;

    const token = normalizeKeywordToken(cleaned);
    if (seen.has(token)) continue;

    seen.add(token);
    result.push(cleaned);
  }

  return result;
}

function parseKeywordInput(value) {
  return normalizeKeywordArray(typeof value === 'string' ? value : '');
}

function normalizeWorldTagMap(value) {
  const output = {};

  if (!value || typeof value !== 'object') {
    return output;
  }

  Object.entries(value).forEach(([worldName, tags]) => {
    if (typeof worldName !== 'string' || !worldName.trim()) return;

    const normalized = normalizeKeywordArray(tags);
    if (normalized.length > 0) {
      output[worldName] = normalized;
    }
  });

  return output;
}

function getAllConfiguredWorldbookTagsInternal(settings) {
  const all = [];
  const seen = new Set();

  Object.values(settings?.worldTags || {}).forEach((tagList) => {
    normalizeKeywordArray(tagList).forEach((tag) => {
      const token = normalizeKeywordToken(tag);
      if (!token || seen.has(token)) return;
      seen.add(token);
      all.push(tag);
    });
  });

  return all.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
}

function getWorldbookManagerSettings() {
  if (!extension_settings.worldInfoSuite) {
    extension_settings.worldInfoSuite = { ...defaultSettings };
  }

  const suite = extension_settings.worldInfoSuite;
  const defaults = createWorldbookManagerDefaults();
  const raw = (suite.worldbookManager && typeof suite.worldbookManager === 'object')
    ? suite.worldbookManager
    : {};

  const normalized = {
    ...defaults,
    ...raw,
    customOrder: Array.isArray(raw.customOrder) ? raw.customOrder.filter((x) => typeof x === 'string') : [],
    priorityKeywords: normalizeKeywordArray(raw.priorityKeywords),
    hiddenKeywords: normalizeKeywordArray(raw.hiddenKeywords),
    prioritizeCharacterBound: Boolean(raw.prioritizeCharacterBound),
    worldTags: normalizeWorldTagMap(raw.worldTags),
    activeTagFilter: typeof raw.activeTagFilter === 'string' ? raw.activeTagFilter.trim() : '',
  };

  if (!Object.values(WORLDBOOK_SORT_MODE).includes(normalized.baseSort)) {
    normalized.baseSort = WORLDBOOK_SORT_MODE.CREATED_DESC;
  }

  if (Array.isArray(world_names) && world_names.length > 0) {
    const validNames = new Set(world_names);
    normalized.customOrder = normalized.customOrder.filter((name) => validNames.has(name));

    Object.keys(normalized.worldTags).forEach((name) => {
      if (!validNames.has(name)) {
        delete normalized.worldTags[name];
      }
    });
  }

  const allTags = getAllConfiguredWorldbookTagsInternal(normalized);
  if (normalized.activeTagFilter && !allTags.some((tag) => normalizeKeywordToken(tag) === normalizeKeywordToken(normalized.activeTagFilter))) {
    normalized.activeTagFilter = '';
  }

  suite.worldbookManager = normalized;
  return normalized;
}

function getWorldbookTags(worldName, settings = getWorldbookManagerSettings()) {
  if (!worldName) return [];
  return normalizeKeywordArray(settings?.worldTags?.[worldName] || []);
}

function getAllConfiguredWorldbookTags() {
  return getAllConfiguredWorldbookTagsInternal(getWorldbookManagerSettings());
}

function worldbookMatchesKeyword(worldName, keyword, settings = getWorldbookManagerSettings()) {
  const needle = normalizeKeywordToken(keyword);
  if (!needle) return false;

  if (normalizeKeywordToken(worldName).includes(needle)) {
    return true;
  }

  return getWorldbookTags(worldName, settings).some((tag) => normalizeKeywordToken(tag).includes(needle));
}

function worldbookHasTag(worldName, tag, settings = getWorldbookManagerSettings()) {
  const needle = normalizeKeywordToken(tag);
  if (!needle) return true;

  return getWorldbookTags(worldName, settings).some((tagName) => normalizeKeywordToken(tagName) === needle);
}

function getCharacterBoundWorldbookSet() {
  const bound = new Set();

  const character = characters?.[this_chid];
  const primaryWorld = character?.data?.extensions?.world;
  if (primaryWorld) {
    bound.add(primaryWorld);
  }

  const fileName = getCharaFilename?.(this_chid);
  const extraCharLore = world_info?.charLore?.find?.((e) => e.name === fileName);
  if (extraCharLore && Array.isArray(extraCharLore.extraBooks)) {
    extraCharLore.extraBooks.forEach((bookName) => {
      if (bookName) bound.add(bookName);
    });
  }

  const chatLoreName = chat_metadata?.[METADATA_KEY];
  if (chatLoreName) {
    bound.add(chatLoreName);
  }

  return bound;
}

function getWorldbookCustomOrder(names, settings) {
  const nameSet = new Set(names);
  const order = settings.customOrder.filter((name) => nameSet.has(name));
  const seen = new Set(order);

  names.forEach((name) => {
    if (!seen.has(name)) {
      order.push(name);
      seen.add(name);
    }
  });

  return order;
}

function sortWorldbookNamesForManager(names, settings = getWorldbookManagerSettings()) {
  const creationIndexMap = new Map(names.map((name, index) => [name, index]));
  const customOrder = getWorldbookCustomOrder(names, settings);
  const customOrderMap = new Map(customOrder.map((name, index) => [name, index]));
  const characterBoundSet = getCharacterBoundWorldbookSet();

  return [...names].sort((a, b) => {
    const aKeywordPriority = settings.priorityKeywords.some((kw) => worldbookMatchesKeyword(a, kw, settings)) ? 1 : 0;
    const bKeywordPriority = settings.priorityKeywords.some((kw) => worldbookMatchesKeyword(b, kw, settings)) ? 1 : 0;
    if (aKeywordPriority !== bKeywordPriority) {
      return bKeywordPriority - aKeywordPriority;
    }

    if (settings.prioritizeCharacterBound) {
      const aBound = characterBoundSet.has(a) ? 1 : 0;
      const bBound = characterBoundSet.has(b) ? 1 : 0;
      if (aBound !== bBound) {
        return bBound - aBound;
      }
    }

    switch (settings.baseSort) {
      case WORLDBOOK_SORT_MODE.CREATED_ASC:
        return (creationIndexMap.get(a) ?? 0) - (creationIndexMap.get(b) ?? 0);
      case WORLDBOOK_SORT_MODE.NAME_ASC:
        return a.localeCompare(b, undefined, { sensitivity: 'base' });
      case WORLDBOOK_SORT_MODE.NAME_DESC:
        return b.localeCompare(a, undefined, { sensitivity: 'base' });
      case WORLDBOOK_SORT_MODE.CUSTOM: {
        const aCustomIndex = customOrderMap.get(a) ?? Number.MAX_SAFE_INTEGER;
        const bCustomIndex = customOrderMap.get(b) ?? Number.MAX_SAFE_INTEGER;
        if (aCustomIndex !== bCustomIndex) {
          return aCustomIndex - bCustomIndex;
        }
        return (creationIndexMap.get(a) ?? 0) - (creationIndexMap.get(b) ?? 0);
      }
      case WORLDBOOK_SORT_MODE.CREATED_DESC:
      default:
        return (creationIndexMap.get(b) ?? 0) - (creationIndexMap.get(a) ?? 0);
    }
  });
}

function getCurrentWorldEditorSelectionName() {
  const sel = /** @type {HTMLSelectElement | null} */ (document.querySelector('#world_editor_select'));
  if (!sel) return '';
  if (sel.value === '') return '';

  const idx = Number(sel.value);
  if (!Number.isNaN(idx) && Array.isArray(world_names) && world_names[idx]) {
    return world_names[idx];
  }

  return '';
}

let worldbookManagerRefreshTimer = null;
let worldbookSelectObserver = null;
let worldbookUiObserver = null;
let worldbookPlaceholderOption;
let isApplyingWorldbookOptions = false;
let worldbookMutationSuppressedUntil = 0;

function rebuildWorldEditorOptions(orderedNames, selectedName) {
  const sel = /** @type {HTMLSelectElement | null} */ (document.querySelector('#world_editor_select'));
  if (!sel || !Array.isArray(world_names)) return;

  if (worldbookPlaceholderOption === undefined) {
    const sourcePlaceholder = sel.querySelector('option[value=""]');
    worldbookPlaceholderOption = sourcePlaceholder ? sourcePlaceholder.cloneNode(true) : false;
    if (worldbookPlaceholderOption) {
      worldbookPlaceholderOption.value = '';
    }
  }

  const previousValue = sel.value;
  const fragment = document.createDocumentFragment();
  if (worldbookPlaceholderOption) {
    fragment.append(worldbookPlaceholderOption.cloneNode(true));
  }

  const worldIndexMap = new Map(world_names.map((name, index) => [name, index]));
  const added = new Set();

  orderedNames.forEach((name) => {
    const idx = worldIndexMap.get(name);
    if (idx === undefined) return;

    const option = new Option(name, String(idx));
    if (name === selectedName) option.selected = true;
    fragment.append(option);
    added.add(name);
  });

  if (selectedName && !added.has(selectedName) && worldIndexMap.has(selectedName)) {
    const selectedOption = new Option(
      `${selectedName} ${i18n('worldManagerHiddenSelectedSuffix')}`,
      String(worldIndexMap.get(selectedName)),
      true,
      true,
    );
    fragment.append(selectedOption);
  }

  worldbookMutationSuppressedUntil = Date.now() + 200;
  isApplyingWorldbookOptions = true;
  sel.replaceChildren(fragment);
  isApplyingWorldbookOptions = false;

  if (selectedName && worldIndexMap.has(selectedName)) {
    const selectedValue = String(worldIndexMap.get(selectedName));
    if (sel.querySelector(`option[value="${selectedValue}"]`)) {
      sel.value = selectedValue;
    }
  }

  if (sel.value === '' && sel.querySelector('option[value=""]')) {
    sel.value = '';
  }

  if (sel.value !== previousValue) {
    sel.dispatchEvent(new Event('change', { bubbles: true }));
  } else {
    $(sel).trigger('change.select2');
  }
}

function applyWorldbookManagerToEditorSelect() {
  const sel = /** @type {HTMLSelectElement | null} */ (document.querySelector('#world_editor_select'));
  if (!sel || !Array.isArray(world_names) || world_names.length === 0) {
    return;
  }

  const selectedName = getCurrentWorldEditorSelectionName();

  if (!extension_settings.worldInfoSuite?.enableWorldbookManager) {
    rebuildWorldEditorOptions([...world_names], selectedName);
    return;
  }

  const settings = getWorldbookManagerSettings();
  const sortedNames = sortWorldbookNamesForManager(world_names, settings);
  const hasActiveTagFilter = normalizeKeywordToken(settings.activeTagFilter) !== '';

  const visibleNames = sortedNames.filter((name) => {
    if (name === selectedName) return true;

    if (!hasActiveTagFilter) {
      const hiddenByKeyword = settings.hiddenKeywords.some((kw) => worldbookMatchesKeyword(name, kw, settings));
      if (hiddenByKeyword) return false;
    }

    if (hasActiveTagFilter && !worldbookHasTag(name, settings.activeTagFilter, settings)) {
      return false;
    }

    return true;
  });

  rebuildWorldEditorOptions(visibleNames, selectedName);
}

function processWorldInfoData(activatedEntries) {
  const byPosition = {};
  const positionInfo = getPositionInfo();
  const selectiveLogicInfo = getSelectiveLogicInfo();

  activatedEntries.forEach((entryRaw) => {
    if (!entryRaw || typeof entryRaw !== 'object') return;

    const position = (typeof entryRaw.position === 'number') ? entryRaw.position : 0;
    const posInfo = positionInfo[position] || { name: `${i18n('positionUnknown')} (${position})`, emoji: '❓' };
    const posKey = `pos_${position}`;

    if (!byPosition[posKey]) {
      byPosition[posKey] = {
        position,
        positionName: posInfo.name,
        positionEmoji: posInfo.emoji,
        entries: [],
      };
    }

    const status = getEntryStatus(entryRaw);
    const sourceKey = getWISourceKey(entryRaw);
    const sourceName = getSourceDisplayName(sourceKey);

    const worldOrder =
      (typeof entryRaw.worldOrder === 'number' ? entryRaw.worldOrder : undefined) ??
      (typeof entryRaw.order === 'number' ? entryRaw.order : undefined) ??
      getWorldOrderByName(entryRaw.world);

    const processedEntry = {
      uid: entryRaw.uid,
      worldName: entryRaw.world,
      entryName: entryRaw.comment || `${i18n('entryLabel')} #${entryRaw.uid}`,
      sourceKey,
      sourceName,
      statusEmoji: status.emoji,
      statusName: status.name,
      content: entryRaw.content,
      keys: Array.isArray(entryRaw.key) ? entryRaw.key.join(', ') : (typeof entryRaw.key === 'string' ? entryRaw.key : null),
      secondaryKeys: Array.isArray(entryRaw.keysecondary) ? entryRaw.keysecondary.join(', ') : null,
      selectiveLogicName: Array.isArray(entryRaw.keysecondary)
        ? (selectiveLogicInfo?.[entryRaw.selectiveLogic] ?? `${i18n('selectiveLogicUnknown')} (${entryRaw.selectiveLogic})`)
        : null,
      selectiveLogic: entryRaw.selectiveLogic ?? null,
      depth: entryRaw.depth ?? null,
      displayDepth: (position === 4) ? (entryRaw.depth ?? null) : null,
      roleDepthTag: (position === 4) ? formatRoleDepthTag(entryRaw) : null,
      role: (entryRaw.role || entryRaw.messageRole || 'assistant'),
      sourceType: getEntrySourceType(entryRaw),
      worldOrder,
    };

    byPosition[posKey].entries.push(processedEntry);
  });

  Object.values(byPosition).forEach((posGroup) => {
    if (posGroup.position === 4) {
      posGroup.entries.sort(compareDepthEntries);
    } else {
      posGroup.entries.sort(compareOrderEntries);
    }
  });

  const groups = Object.values(byPosition).filter(g => g.entries.length > 0);
  groups.sort((a, b) => getPositionSortIndex(a.position) - getPositionSortIndex(b.position));

  return groups;
}

// Re-translate stored world info data to current locale
function retranslateWorldInfoData(worldInfoData) {
  const positionInfo = getPositionInfo();
  const selectiveLogicInfo = getSelectiveLogicInfo();
  const statusEmojiMap = {
    '🔵': () => i18n('statusConstant'),
    '🔗': () => i18n('statusVectorized'),
    '🟢': () => i18n('statusKeyword'),
  };

  return worldInfoData.map(group => {
    const posInfo = positionInfo[group.position] || { name: `${i18n('positionUnknown')} (${group.position})`, emoji: '❓' };
    return {
      ...group,
      positionName: posInfo.name,
      positionEmoji: posInfo.emoji,
      entries: group.entries.map(entry => {
        const statusTranslator = statusEmojiMap[entry.statusEmoji];
        let selectiveLogicName = entry.selectiveLogicName;
        if (entry.secondaryKeys && entry.selectiveLogic != null) {
          selectiveLogicName = selectiveLogicInfo[entry.selectiveLogic]
            ?? `${i18n('selectiveLogicUnknown')} (${entry.selectiveLogic})`;
        }
        return {
          ...entry,
          entryName: entry.entryName || `${i18n('entryLabel')} #${entry.uid}`,
          sourceName: entry.sourceKey ? getSourceDisplayName(entry.sourceKey) : (entry.sourceName || ''),
          statusName: statusTranslator ? statusTranslator() : (entry.statusName || ''),
          roleDepthTag: (group.position === 4 && entry.depth != null) ? formatRoleDepthTag(entry) : entry.roleDepthTag,
          selectiveLogicName,
        };
      }),
    };
  });
}

// ============================================================
// FEATURE 1: Triggered Entry Viewer
// ============================================================

function addViewButtonToMessage(messageId) {
  if (!extension_settings.worldInfoSuite?.enableTriggeredViewer) return;
  if (!chat?.[messageId]?.extra?.worldInfoViewer) return;

  const messageElement = document.querySelector(`.mes[mesid="${messageId}"]`);
  if (!messageElement || messageElement.getAttribute('is_user') === 'true') return;

  // Insert into extraMesButtons instead of mes_buttons to be part of the collapsible menu
  const extraButtonsContainer = messageElement.querySelector('.extraMesButtons');
  if (!extraButtonsContainer) return;

  const buttonId = `worldinfo-viewer-btn-${messageId}`;
  if (document.getElementById(buttonId)) return;

  const iconClass = extension_settings.worldInfoSuite?.viewerIcon || 'fa-globe';
  const button = document.createElement('div');
  button.id = buttonId;
  button.className = `mes_button worldinfo-viewer-btn fa-regular ${iconClass}`;
  button.title = i18n('viewerBtnTitle');
  button.addEventListener('click', (event) => {
    event.stopPropagation();
    showWorldInfoPopup(messageId);
  });

  // Insert at the beginning of extraMesButtons
  extraButtonsContainer.prepend(button);
}

async function showWorldInfoPopup(messageId) {
  const worldInfoData = chat?.[messageId]?.extra?.worldInfoViewer;
  if (!worldInfoData) {
    toastr.info(i18n('noWorldInfoData'));
    return;
  }

  // Re-translate stored data to current locale before display
  const translatedData = retranslateWorldInfoData(worldInfoData);

  try {
    const popupContent = await renderExtensionTemplateAsync(extensionName, 'popup', {
      positions: translatedData,
      i18n: localeData,
    });
    callGenericPopup(popupContent, POPUP_TYPE.TEXT, '', {
      wide: true,
      large: true,
      okButton: i18n('popupClose'),
      allowVerticalScrolling: true,
    });
  } catch (error) {
    console.error(`[${extensionName}] Error rendering popup:`, error);
    toastr.error(i18n('popupRenderError'));
  }
}

// State sync for triggered viewer
let lastActivatedWorldInfo = null;

// Clean up old World Info viewer data to limit cache size
function cleanupViewerCache() {
  if (!chat || !Array.isArray(chat)) return;
  
  const limit = extension_settings.worldInfoSuite?.viewerCacheLimit ?? 10;
  if (limit <= 0) return; // 0 means unlimited
  
  // Find all messages with viewer data
  const messagesWithData = [];
  for (let i = 0; i < chat.length; i++) {
    if (chat[i]?.extra?.worldInfoViewer) {
      messagesWithData.push(i);
    }
  }
  
  // Remove data from older messages beyond the limit
  if (messagesWithData.length > limit) {
    const toRemove = messagesWithData.slice(0, messagesWithData.length - limit);
    for (const idx of toRemove) {
      if (chat[idx]?.extra?.worldInfoViewer) {
        delete chat[idx].extra.worldInfoViewer;
      }
    }
  }
}

// Clear all World Info viewer cache from current chat
function clearAllViewerCache() {
  if (!chat || !Array.isArray(chat)) return 0;
  
  let count = 0;
  for (let i = 0; i < chat.length; i++) {
    if (chat[i]?.extra?.worldInfoViewer) {
      delete chat[i].extra.worldInfoViewer;
      count++;
    }
  }
  
  // Remove viewer buttons from UI
  document.querySelectorAll('.worldinfo-viewer-btn').forEach(btn => btn.remove());
  
  return count;
}

function initTriggeredViewer() {
  eventSource.on(event_types.WORLD_INFO_ACTIVATED, (data) => {
    if (!extension_settings.worldInfoSuite?.enableTriggeredViewer) return;
    if (data && Array.isArray(data) && data.length > 0) {
      lastActivatedWorldInfo = processWorldInfoData(data);
    } else {
      lastActivatedWorldInfo = null;
    }
  });

  eventSource.on(event_types.MESSAGE_RECEIVED, (messageId) => {
    if (!extension_settings.worldInfoSuite?.enableTriggeredViewer) return;
    if (lastActivatedWorldInfo && chat?.[messageId] && !chat[messageId].is_user) {
      if (!chat[messageId].extra) chat[messageId].extra = {};
      chat[messageId].extra.worldInfoViewer = lastActivatedWorldInfo;
      lastActivatedWorldInfo = null;
      
      // Clean up old cache after adding new data
      cleanupViewerCache();
    }
  });

  eventSource.on(event_types.CHARACTER_MESSAGE_RENDERED, (messageId) => {
    addViewButtonToMessage(String(messageId));
  });

  eventSource.on(event_types.CHAT_CHANGED, () => {
    setTimeout(() => {
      document.querySelectorAll('#chat .mes').forEach((messageElement) => {
        const mesId = messageElement.getAttribute('mesid');
        if (mesId) addViewButtonToMessage(mesId);
      });
    }, 500);
  });
}

// ============================================================
// FEATURE 2: Character Lorebook Quick Access
// ============================================================

function isMobileDevice() {
  return window.innerWidth <= 768;
}

function shouldShowGlobalLorebooks() {
  const settings = extension_settings.worldInfoSuite;
  if (isMobileDevice()) {
    return settings?.showGlobalLorebookMobile ?? true;
  }
  return settings?.showGlobalLorebookDesktop ?? true;
}

function getCharacterWorldBooks(chid) {
  const books = [];
  const character = characters?.[chid];

  if (!character) return books;

  // Primary Lorebook
  const primaryWorld = character?.data?.extensions?.world;
  if (primaryWorld && world_names?.includes(primaryWorld)) {
    books.push({ name: primaryWorld, type: 'primary' });
  }

  // Additional Lorebooks
  const fileName = getCharaFilename?.(chid);
  const extraCharLore = world_info?.charLore?.find?.((e) => e.name === fileName);
  if (extraCharLore && Array.isArray(extraCharLore.extraBooks)) {
    extraCharLore.extraBooks.forEach((bookName) => {
      if (bookName && world_names?.includes(bookName)) {
        books.push({ name: bookName, type: 'additional' });
      }
    });
  }

  // Chat Lorebook
  const chatLoreName = chat_metadata?.[METADATA_KEY];
  if (chatLoreName && world_names?.includes(chatLoreName)) {
    books.push({ name: chatLoreName, type: 'chat' });
  }

  // Global Lorebooks (always at the end)
  if (shouldShowGlobalLorebooks() && Array.isArray(selected_world_info)) {
    selected_world_info.forEach((worldName) => {
      if (worldName && world_names?.includes(worldName)) {
        books.push({ name: worldName, type: 'global' });
      }
    });
  }

  return books;
}

function createCharacterWorldBooksHTML(books) {
  if (books.length === 0) {
    return `
      <div id="char-worldbooks-panel" class="char-worldbooks-panel">
        <div class="char-worldbooks-header">
          <span class="fa-solid fa-globe"></span>
          <span data-i18n="charWorldbooksPanel">${i18n('charWorldbooksPanel')}</span>
        </div>
        <div class="char-worldbooks-empty" data-i18n="charWorldbooksEmpty">
          ${i18n('charWorldbooksEmpty')}
        </div>
      </div>
    `;
  }

  const bookItems = books.map((book) => {
    let typeLabel, typeClass;
    if (book.type === 'primary') {
      typeLabel = i18n('charWorldbookTypePrimary');
      typeClass = 'primary';
    } else if (book.type === 'chat') {
      typeLabel = i18n('charWorldbookTypeChat');
      typeClass = 'chat';
    } else if (book.type === 'global') {
      typeLabel = i18n('charWorldbookTypeGlobal');
      typeClass = 'global';
    } else {
      typeLabel = i18n('charWorldbookTypeAdditional');
      typeClass = 'additional';
    }
    return `
      <div class="char-worldbook-item" data-world-name="${book.name}">
        <span class="char-worldbook-type ${typeClass}">${typeLabel}</span>
        <span class="char-worldbook-name">${book.name}</span>
        <span class="char-worldbook-goto fa-solid fa-arrow-up-right-from-square" title="${i18n('charWorldbookGotoTitle')}"></span>
      </div>
    `;
  }).join('');

  return `
    <div id="char-worldbooks-panel" class="char-worldbooks-panel">
      <div class="char-worldbooks-header">
        <span class="fa-solid fa-globe"></span>
        <span data-i18n="charWorldbooksPanel">${i18n('charWorldbooksPanel')}</span>
      </div>
      <div class="char-worldbooks-list">
        ${bookItems}
      </div>
    </div>
  `;
}

function updateCharacterWorldBooksPanel(chid) {
  if (!extension_settings.worldInfoSuite?.enableCharLorebook) return;

  $('#char-worldbooks-panel').remove();

  const books = getCharacterWorldBooks(chid);
  const panelHTML = createCharacterWorldBooksHTML(books);

  const dropdownLabel = $('#char-management-dropdown').closest('label');
  if (dropdownLabel.length) {
    dropdownLabel.after(panelHTML);
  } else {
    $('#avatar_controls').append(panelHTML);
  }

  $('.char-worldbook-item').on('click', function () {
    const worldName = $(this).data('world-name');
    if (worldName && typeof openWorldInfoEditor === 'function') {
      openWorldInfoEditor(worldName);
    }
  });
}

function hideCharacterWorldBooksPanel() {
  $('#char-worldbooks-panel').remove();
}

function initCharLorebookQuickAccess() {
  // Store the currently opened chid for use in other event handlers
  let currentEditorChid = null;

  eventSource.on(event_types.CHARACTER_EDITOR_OPENED, (chid) => {
    currentEditorChid = chid;
    if (extension_settings.worldInfoSuite?.enableCharLorebook) {
      updateCharacterWorldBooksPanel(chid);
    }
  });

  eventSource.on(event_types.CHARACTER_EDITED, (data) => {
    const chid = data?.detail?.id;
    if (chid !== undefined && $('#char-worldbooks-panel').length) {
      currentEditorChid = chid;
      if (extension_settings.worldInfoSuite?.enableCharLorebook) {
        updateCharacterWorldBooksPanel(chid);
      }
    }
  });

  // Listen for World Info settings updates (when lorebooks are bound/unbound)
  eventSource.on(event_types.WORLDINFO_SETTINGS_UPDATED, () => {
    if (currentEditorChid !== null && $('#char-worldbooks-panel').length) {
      if (extension_settings.worldInfoSuite?.enableCharLorebook) {
        updateCharacterWorldBooksPanel(currentEditorChid);
      }
    }
  });

  // Listen for chat changed event to update chat lorebook
  eventSource.on(event_types.CHAT_CHANGED, () => {
    if (currentEditorChid !== null && $('#char-worldbooks-panel').length) {
      if (extension_settings.worldInfoSuite?.enableCharLorebook) {
        updateCharacterWorldBooksPanel(currentEditorChid);
      }
    }
  });

  // Use MutationObserver to detect popup close events for world info binding changes
  // This is needed because ST doesn't emit events when auxiliary lorebooks are changed
  const popupObserver = new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      // Check for removed popup dialogs
      for (const node of mutation.removedNodes) {
        if (node instanceof HTMLElement && node.classList?.contains('popup')) {
          // A popup was closed, check if we need to update the panel
          if (currentEditorChid !== null && $('#char-worldbooks-panel').length) {
            if (extension_settings.worldInfoSuite?.enableCharLorebook) {
              // Small delay to ensure settings are saved
              setTimeout(() => {
                updateCharacterWorldBooksPanel(currentEditorChid);
              }, 100);
            }
          }
        }
      }
    }
  });

  // Observe the document body for popup removals
  popupObserver.observe(document.body, { childList: true });
}

// ============================================================
// FEATURE 3: Worldbook Sort & Keyword Manager
// ============================================================

function renderWorldbookQuickTagBar(select2Container = null) {
  const sel = /** @type {HTMLSelectElement | null} */ (document.querySelector('#world_editor_select'));
  if (!sel) return;

  const container = select2Container || sel.nextElementSibling;
  if (!(container instanceof HTMLElement) || !container.classList.contains('select2')) return;

  const controlRow = container.closest('.flex-container.alignitemscenter');
  const host = container.parentElement || container;
  let tagBar = document.querySelector('.wis-worldbook-tag-bar');
  if (!tagBar) {
    tagBar = document.createElement('div');
    tagBar.classList.add('wis-worldbook-tag-bar');
  }

  if (controlRow instanceof HTMLElement) {
    if (tagBar.previousElementSibling !== controlRow || tagBar.parentElement !== controlRow.parentElement) {
      controlRow.insertAdjacentElement('afterend', tagBar);
    }
  } else if (tagBar.parentElement !== host) {
    container.insertAdjacentElement('afterend', tagBar);
  }

  if (!extension_settings.worldInfoSuite?.enableWorldbookManager) {
    tagBar.innerHTML = '';
    tagBar.style.display = 'none';
    return;
  }

  const settings = getWorldbookManagerSettings();
  const tags = getAllConfiguredWorldbookTagsInternal(settings);

  if (tags.length === 0) {
    tagBar.innerHTML = '';
    tagBar.style.display = 'none';
    return;
  }

  tagBar.style.display = 'flex';
  tagBar.innerHTML = '';

  const label = document.createElement('span');
  label.classList.add('wis-worldbook-tag-label');
  label.textContent = i18n('worldManagerQuickFilterLabel');
  tagBar.append(label);

  const activeTagToken = normalizeKeywordToken(settings.activeTagFilter);

  const makeChip = (displayText, tagValue, isActive) => {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.classList.add('menu_button', 'menu_button_small', 'wis-worldbook-tag-chip');
    if (isActive) {
      chip.classList.add('active');
    }
    chip.textContent = displayText;
    chip.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      const latest = getWorldbookManagerSettings();
      latest.activeTagFilter = tagValue;
      saveSettingsDebounced();
      scheduleWorldbookManagerRefresh();
    });
    tagBar.append(chip);
  };

  makeChip(i18n('worldManagerQuickFilterAll'), '', activeTagToken === '');

  tags.forEach((tag) => {
    const token = normalizeKeywordToken(tag);
    makeChip(tag, tag, token === activeTagToken);
  });
}

function ensureWorldbookManagerControls() {
  const sel = /** @type {HTMLSelectElement | null} */ (document.querySelector('#world_editor_select'));
  if (!sel) return false;

  const select2Container = sel.nextElementSibling;
  if (!(select2Container instanceof HTMLElement) || !select2Container.classList.contains('select2')) {
    return false;
  }

  select2Container.classList.add('wis-worldbook-select2');

  const row = select2Container.parentElement;
  if (!(row instanceof HTMLElement)) {
    return false;
  }

  let manageBtn = row.querySelector('.wis-worldbook-manage-btn');
  if (!(manageBtn instanceof HTMLElement)) {
    manageBtn = document.createElement('div');
    manageBtn.classList.add('menu_button', 'fa-solid', 'fa-arrow-down-a-z', 'interactable', 'wis-worldbook-manage-btn');
    manageBtn.setAttribute('role', 'button');
    manageBtn.tabIndex = 0;

    const openManager = (event) => {
      event.preventDefault();
      event.stopPropagation();

      if (!extension_settings.worldInfoSuite?.enableWorldbookManager) {
        toastr.warning(i18n('worldManagerDisabled'));
        return;
      }

      showWorldbookManagerDialog();
    };

    manageBtn.addEventListener('click', openManager);
    manageBtn.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        openManager(event);
      }
    });
  }

  if (manageBtn.parentElement !== row || manageBtn.nextElementSibling !== select2Container) {
    row.insertBefore(manageBtn, select2Container);
  }

  manageBtn.title = i18n('worldManagerBtnTitle');
  manageBtn.setAttribute('aria-label', i18n('worldManagerBtnTitle'));
  manageBtn.style.display = extension_settings.worldInfoSuite?.enableWorldbookManager ? '' : 'none';

  renderWorldbookQuickTagBar(select2Container);
  return true;
}

async function showWorldbookManagerDialog() {
  if (!Array.isArray(world_names) || world_names.length === 0) {
    toastr.warning(i18n('worldManagerNoWorlds'));
    return;
  }

  const settings = getWorldbookManagerSettings();
  const names = [...world_names];
  const customOrder = getWorldbookCustomOrder(names, settings);

  const dom = document.createElement('div');
  dom.classList.add('wis-worldbook-manager-dialog');

  const title = document.createElement('h3');
  title.textContent = i18n('worldManagerPopupTitle');
  dom.append(title);

  const hint = document.createElement('small');
  hint.classList.add('wis-worldbook-manager-hint');
  hint.textContent = i18n('worldManagerPopupHint');
  dom.append(hint);

  const createSection = (labelKey) => {
    const section = document.createElement('section');
    section.classList.add('wis-worldbook-manager-section');
    const sectionTitle = document.createElement('h4');
    sectionTitle.textContent = i18n(labelKey);
    section.append(sectionTitle);
    dom.append(section);
    return section;
  };

  let selectedSort = settings.baseSort;

  const simpleSortSection = createSection('worldManagerSectionSimpleSort');
  const sortButtonGroup = document.createElement('div');
  sortButtonGroup.classList.add('wis-worldbook-sort-buttons');
  simpleSortSection.append(sortButtonGroup);

  const sortOptions = [
    { mode: WORLDBOOK_SORT_MODE.CREATED_DESC, key: 'worldManagerSortCreatedDesc' },
    { mode: WORLDBOOK_SORT_MODE.CREATED_ASC, key: 'worldManagerSortCreatedAsc' },
    { mode: WORLDBOOK_SORT_MODE.NAME_ASC, key: 'worldManagerSortNameAsc' },
    { mode: WORLDBOOK_SORT_MODE.NAME_DESC, key: 'worldManagerSortNameDesc' },
    { mode: WORLDBOOK_SORT_MODE.CUSTOM, key: 'worldManagerSortCustom' },
  ];

  const sortButtons = [];
  const refreshSortButtonState = () => {
    sortButtons.forEach(({ button, mode }) => {
      button.classList.toggle('active', mode === selectedSort);
    });
  };

  const updateCustomSectionVisibility = (section) => {
    section.style.display = selectedSort === WORLDBOOK_SORT_MODE.CUSTOM ? '' : 'none';
  };

  sortOptions.forEach(({ mode, key }) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.classList.add('menu_button', 'menu_button_small', 'wis-worldbook-sort-btn');
    button.textContent = i18n(key);
    button.addEventListener('click', () => {
      selectedSort = mode;
      refreshSortButtonState();
      updateCustomSectionVisibility(customSection);
    });
    sortButtonGroup.append(button);
    sortButtons.push({ button, mode });
  });

  const prioritySection = createSection('worldManagerSectionPriority');

  const priorityLabel = document.createElement('label');
  priorityLabel.classList.add('wis-worldbook-manager-input-label');
  priorityLabel.textContent = i18n('worldManagerPriorityKeywords');
  prioritySection.append(priorityLabel);

  const priorityInput = document.createElement('input');
  priorityInput.type = 'text';
  priorityInput.classList.add('text_pole');
  priorityInput.placeholder = i18n('worldManagerPriorityKeywordsPlaceholder');
  priorityInput.value = settings.priorityKeywords.join(', ');
  prioritySection.append(priorityInput);

  const boundPriorityLabel = document.createElement('label');
  boundPriorityLabel.classList.add('wis-feature-label', 'wis-worldbook-inline-checkbox');
  const boundPriorityCheckbox = document.createElement('input');
  boundPriorityCheckbox.type = 'checkbox';
  boundPriorityCheckbox.checked = Boolean(settings.prioritizeCharacterBound);
  const boundPriorityText = document.createElement('span');
  boundPriorityText.textContent = i18n('worldManagerPrioritizeBound');
  boundPriorityLabel.append(boundPriorityCheckbox, boundPriorityText);
  prioritySection.append(boundPriorityLabel);

  const hideSection = createSection('worldManagerSectionHide');

  const hiddenLabel = document.createElement('label');
  hiddenLabel.classList.add('wis-worldbook-manager-input-label');
  hiddenLabel.textContent = i18n('worldManagerHiddenKeywords');
  hideSection.append(hiddenLabel);

  const hiddenInput = document.createElement('input');
  hiddenInput.type = 'text';
  hiddenInput.classList.add('text_pole');
  hiddenInput.placeholder = i18n('worldManagerHiddenKeywordsPlaceholder');
  hiddenInput.value = settings.hiddenKeywords.join(', ');
  hideSection.append(hiddenInput);

  const customSection = createSection('worldManagerSectionCustomOrder');
  const customHint = document.createElement('small');
  customHint.classList.add('wis-worldbook-manager-hint');
  customHint.textContent = i18n('worldManagerCustomOrderHint');
  customSection.append(customHint);

  const customOrderList = document.createElement('div');
  customOrderList.classList.add('wis-worldbook-order-list');
  customSection.append(customOrderList);

  refreshSortButtonState();
  updateCustomSectionVisibility(customSection);

  customOrder.forEach((worldName) => {
    const item = document.createElement('div');
    item.classList.add('wis-worldbook-order-item');
    item.dataset.worldName = worldName;

    const handle = document.createElement('span');
    handle.classList.add('wis-worldbook-order-handle', 'fa-solid', 'fa-grip-lines');
    const nameEl = document.createElement('span');
    nameEl.classList.add('wis-worldbook-order-name');
    nameEl.textContent = worldName;

    item.append(handle, nameEl);
    customOrderList.append(item);
  });

  let sortableEnabled = false;
  if (typeof $(customOrderList).sortable === 'function') {
    sortableEnabled = true;
    $(customOrderList).sortable({
      items: '.wis-worldbook-order-item',
      handle: '.wis-worldbook-order-handle',
      axis: 'y',
      tolerance: 'pointer',
    });
  }

  const tagSection = createSection('worldManagerSectionTagEditor');

  const tagSearchInput = document.createElement('input');
  tagSearchInput.type = 'text';
  tagSearchInput.classList.add('text_pole');
  tagSearchInput.placeholder = i18n('worldManagerTagSearchPlaceholder');
  tagSection.append(tagSearchInput);

  const tagRowsWrap = document.createElement('div');
  tagRowsWrap.classList.add('wis-worldbook-tag-editor-list');
  tagSection.append(tagRowsWrap);

  const tagRows = [];

  customOrder.forEach((worldName) => {
    const row = document.createElement('label');
    row.classList.add('wis-worldbook-tag-editor-row');
    row.dataset.worldName = worldName;

    const nameEl = document.createElement('span');
    nameEl.classList.add('wis-worldbook-tag-editor-name');
    nameEl.textContent = worldName;

    const input = document.createElement('input');
    input.type = 'text';
    input.classList.add('text_pole');
    input.placeholder = i18n('worldManagerTagInputPlaceholder');
    input.value = getWorldbookTags(worldName, settings).join(', ');

    row.append(nameEl, input);
    tagRowsWrap.append(row);
    tagRows.push({ row, worldName, input });
  });

  tagSearchInput.addEventListener('input', () => {
    const token = normalizeKeywordToken(tagSearchInput.value);

    tagRows.forEach(({ row, worldName, input }) => {
      if (!token) {
        row.style.display = '';
        return;
      }

      const inName = normalizeKeywordToken(worldName).includes(token);
      const inTags = normalizeKeywordToken(input.value).includes(token);
      row.style.display = (inName || inTags) ? '' : 'none';
    });
  });

  const popup = new Popup(dom, POPUP_TYPE.CONFIRM, null, {
    okButton: i18n('worldManagerApply'),
    cancelButton: i18n('bulkEditCancel'),
    large: true,
    wider: true,
    allowVerticalScrolling: true,
  });

  const result = await popup.show();

  if (sortableEnabled) {
    try {
      if ($(customOrderList).sortable('instance') !== undefined) {
        $(customOrderList).sortable('destroy');
      }
    } catch (_err) {
      // ignore cleanup errors
    }
  }

  if (result !== POPUP_RESULT.AFFIRMATIVE) {
    return;
  }

  settings.baseSort = selectedSort;
  settings.priorityKeywords = parseKeywordInput(priorityInput.value);
  settings.hiddenKeywords = parseKeywordInput(hiddenInput.value);
  settings.prioritizeCharacterBound = boundPriorityCheckbox.checked;
  settings.customOrder = [...customOrderList.querySelectorAll('.wis-worldbook-order-item')]
    .map((item) => item.dataset.worldName)
    .filter(Boolean);

  const worldTags = {};
  tagRows.forEach(({ worldName, input }) => {
    const tags = parseKeywordInput(input.value);
    if (tags.length > 0) {
      worldTags[worldName] = tags;
    }
  });
  settings.worldTags = worldTags;

  const availableTags = getAllConfiguredWorldbookTagsInternal(settings);
  if (settings.activeTagFilter && !availableTags.some((tag) => normalizeKeywordToken(tag) === normalizeKeywordToken(settings.activeTagFilter))) {
    settings.activeTagFilter = '';
  }

  saveSettingsDebounced();
  scheduleWorldbookManagerRefresh();
  toastr.success(i18n('worldManagerSaved'));
}

function refreshWorldbookManagerUI() {
  const controlsReady = ensureWorldbookManagerControls();
  if (!controlsReady) return;

  applyWorldbookManagerToEditorSelect();
  renderWorldbookQuickTagBar();
}

function scheduleWorldbookManagerRefresh() {
  if (worldbookManagerRefreshTimer) {
    clearTimeout(worldbookManagerRefreshTimer);
  }

  worldbookManagerRefreshTimer = setTimeout(() => {
    worldbookManagerRefreshTimer = null;
    refreshWorldbookManagerUI();
  }, 80);
}

function initWorldbookManager() {
  const bindSelectObserver = () => {
    const sel = document.querySelector('#world_editor_select');
    if (!sel) return;

    if (worldbookSelectObserver) {
      worldbookSelectObserver.disconnect();
    }

    worldbookSelectObserver = new MutationObserver(() => {
      if (isApplyingWorldbookOptions) return;
      if (Date.now() < worldbookMutationSuppressedUntil) return;
      scheduleWorldbookManagerRefresh();
    });

    worldbookSelectObserver.observe(sel, { childList: true });
  };

  const trySetup = () => {
    const ok = ensureWorldbookManagerControls();
    if (!ok) return false;

    bindSelectObserver();
    scheduleWorldbookManagerRefresh();
    return true;
  };

  if (!trySetup()) {
    let retries = 0;
    const retryTimer = setInterval(() => {
      retries += 1;
      if (trySetup() || retries > 40) {
        clearInterval(retryTimer);
      }
    }, 250);
  }

  if (worldbookUiObserver) {
    worldbookUiObserver.disconnect();
  }

  worldbookUiObserver = new MutationObserver(() => {
    if (Date.now() < worldbookMutationSuppressedUntil) return;

    const select2Container = document.querySelector('#world_editor_select + .select2');
    const row = select2Container?.parentElement;
    const hasManageBtn = row instanceof HTMLElement
      && Array.from(row.children).some((child) => child instanceof HTMLElement && child.classList.contains('wis-worldbook-manage-btn'));

    if (select2Container && !hasManageBtn) {
      scheduleWorldbookManagerRefresh();
    }
  });

  worldbookUiObserver.observe(document.body, { childList: true, subtree: true });

  eventSource.on(event_types.WORLDINFO_SETTINGS_UPDATED, () => {
    scheduleWorldbookManagerRefresh();
  });

  eventSource.on(event_types.CHAT_CHANGED, () => {
    scheduleWorldbookManagerRefresh();
  });
}

function updateWorldbookManagerVisibility() {
  const enabled = extension_settings.worldInfoSuite?.enableWorldbookManager;

  const manageBtn = document.querySelector('.wis-worldbook-manage-btn');
  if (manageBtn instanceof HTMLElement) {
    manageBtn.style.display = enabled ? '' : 'none';
  }

  const tagBar = document.querySelector('.wis-worldbook-tag-bar');
  if (tagBar instanceof HTMLElement) {
    tagBar.style.display = enabled ? '' : 'none';
  }

  scheduleWorldbookManagerRefresh();
}

// ============================================================
// FEATURE 4: Bulk Entry Editor
// ============================================================

async function copyTextToClipboard(text) {
  const normalized = String(text ?? '');

  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(normalized);
      return true;
    }
  } catch (_err) {
    // Fallback below
  }

  try {
    const textArea = document.createElement('textarea');
    textArea.value = normalized;
    textArea.setAttribute('readonly', 'true');
    textArea.style.position = 'fixed';
    textArea.style.top = '-9999px';
    textArea.style.left = '-9999px';
    document.body.append(textArea);
    textArea.select();
    textArea.setSelectionRange(0, textArea.value.length);
    const copied = document.execCommand('copy');
    textArea.remove();
    return copied;
  } catch (_err) {
    return false;
  }
}

function getEntryDisplayTitle(entry) {
  if (!entry || typeof entry !== 'object') {
    return i18n('entryLabel');
  }

  if (entry.comment && String(entry.comment).trim()) {
    return String(entry.comment).trim();
  }

  if (Array.isArray(entry.key) && entry.key.length > 0) {
    return entry.key.join(', ');
  }

  return `${i18n('entryLabel')} #${entry.uid ?? ''}`.trim();
}

async function copyEntryContentsByUid(entriesByUid, uids) {
  const orderedUids = Array.isArray(uids) ? uids : [];
  const entries = orderedUids
    .map((uid) => entriesByUid?.[uid])
    .filter((entry) => entry && typeof entry === 'object');

  if (entries.length === 0) {
    toastr.warning(i18n('bulkCopyEntriesNoSelection'));
    return false;
  }

  const payload = entries
    .map((entry) => String(entry.content ?? ''))
    .join('\n\n');

  const copied = await copyTextToClipboard(payload);
  if (!copied) {
    toastr.error(i18n('bulkCopyEntriesClipboardFailed'));
    return false;
  }

  if (entries.length === 1) {
    toastr.success(i18n('bulkCopyEntriesSingleSuccess', getEntryDisplayTitle(entries[0])));
  } else {
    toastr.success(i18n('bulkCopyEntriesMultiSuccess', entries.length));
  }

  return true;
}

let nativeEntryCopyObserver = null;
let nativeEntryCopyMountTimer = null;
const NATIVE_ENTRY_COPY_RETRY_DELAY = 45;
const NATIVE_ENTRY_COPY_MAX_RETRY = 12;

function getEntryDisplayTitleFromForm(form) {
  if (!(form instanceof HTMLElement)) {
    return i18n('entryLabel');
  }

  const commentInput = form.querySelector('[name="comment"]');
  if (commentInput && 'value' in commentInput) {
    const comment = String(commentInput.value ?? '').trim();
    if (comment) {
      return comment;
    }
  }

  const worldEntry = form.closest('.world_entry');
  const uid = worldEntry?.getAttribute('uid') || worldEntry?.getAttribute('data-uid') || '';
  return `${i18n('entryLabel')} #${uid}`.trim();
}

function mountNativeEntrySingleCopyButton(form) {
  if (!(form instanceof HTMLElement)) {
    return false;
  }

  const contentRow = form.querySelector('span.alignitemscenter.flex-container.flexnowrap.wide100p.justifySpaceBetween');
  if (!(contentRow instanceof HTMLElement)) {
    return false;
  }

  const leftGroup = contentRow.querySelector('span.alignitemscenter.flex-container')
    || contentRow.firstElementChild;
  if (!(leftGroup instanceof HTMLElement)) {
    return false;
  }

  let copyBtn = leftGroup.querySelector('.wis-entry-content-inline-copy-btn');
  if (!(copyBtn instanceof HTMLButtonElement)) {
    copyBtn = document.createElement('button');
    copyBtn.type = 'button';
    copyBtn.classList.add('menu_button', 'menu_button_small', 'fa-regular', 'fa-copy', 'wis-entry-content-inline-copy-btn');
    copyBtn.addEventListener('click', async (event) => {
      event.preventDefault();
      event.stopPropagation();

      const contentInput = form.querySelector('[name="content"]');
      const text = (contentInput && 'value' in contentInput)
        ? String(contentInput.value ?? '')
        : '';

      const copied = await copyTextToClipboard(text);
      if (!copied) {
        toastr.error(i18n('bulkCopyEntriesClipboardFailed'));
        return;
      }

      toastr.success(i18n('bulkCopyEntriesSingleSuccess', getEntryDisplayTitleFromForm(form)));
    });
  }

  copyBtn.title = i18n('bulkCopyEntriesSingleButtonTitle');
  copyBtn.style.display = extension_settings.worldInfoSuite?.enableBulkEditor ? '' : 'none';

  const expandBtn = leftGroup.querySelector('.editor_maximize');
  if (expandBtn instanceof HTMLElement) {
    if (expandBtn.nextElementSibling !== copyBtn) {
      expandBtn.insertAdjacentElement('afterend', copyBtn);
    }
  } else if (copyBtn.parentElement !== leftGroup) {
    leftGroup.append(copyBtn);
  }

  return true;
}

function scheduleNativeEntryCopyMountForForm(form, attempt = 0) {
  if (!(form instanceof HTMLElement)) {
    return;
  }

  if (attempt === 0) {
    if (form.dataset.wisCopyMountPending === '1') {
      return;
    }
    form.dataset.wisCopyMountPending = '1';
  }

  const delayMs = attempt === 0 ? 0 : NATIVE_ENTRY_COPY_RETRY_DELAY;
  setTimeout(() => {
    const mounted = mountNativeEntrySingleCopyButton(form);
    if (!mounted && attempt < NATIVE_ENTRY_COPY_MAX_RETRY) {
      scheduleNativeEntryCopyMountForForm(form, attempt + 1);
      return;
    }

    delete form.dataset.wisCopyMountPending;
  }, delayMs);
}

function updateNativeEntryCopyButtonsVisibility() {
  const enabled = extension_settings.worldInfoSuite?.enableBulkEditor;
  document.querySelectorAll('#world_popup_entries_list .wis-entry-content-inline-copy-btn').forEach((btn) => {
    if (btn instanceof HTMLElement) {
      btn.style.display = enabled ? '' : 'none';
    }
  });
}

function mountNativeEntryCopyButtons(root = document) {
  if (!(root instanceof Document) && !(root instanceof HTMLElement)) {
    return;
  }

  root.querySelectorAll('#world_popup_entries_list .world_entry form.world_entry_form').forEach((form) => {
    scheduleNativeEntryCopyMountForForm(form);
  });

  updateNativeEntryCopyButtonsVisibility();
}

function scheduleNativeEntryCopyButtonsMount() {
  if (nativeEntryCopyMountTimer) {
    clearTimeout(nativeEntryCopyMountTimer);
  }

  nativeEntryCopyMountTimer = setTimeout(() => {
    nativeEntryCopyMountTimer = null;
    mountNativeEntryCopyButtons(document);
  }, 16);
}

function initNativeEntryCopyButtons() {
  const bindObserver = () => {
    const entryList = document.querySelector('#world_popup_entries_list');
    if (!(entryList instanceof HTMLElement)) {
      return false;
    }

    if (nativeEntryCopyObserver) {
      nativeEntryCopyObserver.disconnect();
    }

    nativeEntryCopyObserver = new MutationObserver((mutations) => {
      const shouldMount = mutations.some((mutation) => mutation.type === 'childList' && (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0));
      if (shouldMount) {
        scheduleNativeEntryCopyButtonsMount();
      }
    });

    nativeEntryCopyObserver.observe(entryList, { childList: true, subtree: true });

    entryList.addEventListener('click', (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        return;
      }

      const toggle = target.closest('.inline-drawer-toggle');
      if (!(toggle instanceof HTMLElement)) {
        return;
      }

      const form = toggle.closest('form.world_entry_form');
      if (!(form instanceof HTMLElement)) {
        return;
      }

      scheduleNativeEntryCopyMountForForm(form);
    }, true);

    scheduleNativeEntryCopyButtonsMount();
    return true;
  };

  if (!bindObserver()) {
    let retries = 0;
    const retryTimer = setInterval(() => {
      retries += 1;
      if (bindObserver() || retries > 40) {
        clearInterval(retryTimer);
      }
    }, 250);
  }

  eventSource.on(event_types.WORLDINFO_SETTINGS_UPDATED, () => {
    scheduleNativeEntryCopyButtonsMount();
  });

  eventSource.on(event_types.CHAT_CHANGED, () => {
    scheduleNativeEntryCopyButtonsMount();
  });
}

function initBulkEditor() {
  const btn = document.createElement('div');
  btn.id = 'wis-bulk-edit-btn';
  btn.classList.add('wis-bulk-trigger', 'menu_button', 'fa-solid', 'fa-list-check');
  btn.title = i18n('bulkEditBtnTitle');

  btn.addEventListener('click', async () => {
    if (!extension_settings.worldInfoSuite?.enableBulkEditor) {
      toastr.warning(i18n('bulkEditDisabled'));
      return;
    }

    const sel = /** @type {HTMLSelectElement} */ (document.querySelector('#world_editor_select'));
    if (!sel || /** @type {HTMLOptionElement} */ (sel.children[0])?.selected) {
      toastr.warning(i18n('bulkEditNoWorldSelected'));
      return;
    }

    const selectedIndex = Number(sel.value);
    const name = (!Number.isNaN(selectedIndex) && Array.isArray(world_names) && world_names[selectedIndex])
      ? world_names[selectedIndex]
      : sel.selectedOptions[0].textContent;
    const data = await loadWorldInfo(name);

    // Create a template entry
    const entry = createWorldInfoEntry(null, data);
    await saveWorldInfo(name, data, true);
    sel.dispatchEvent(new Event('change', { bubbles: true }));

    // Wait for the template form to render
    let form = document.querySelector(`#world_popup_entries_list .world_entry[uid="${entry.uid}"]`);
    while (!form) {
      await delay(100);
      form = document.querySelector(`#world_popup_entries_list .world_entry[uid="${entry.uid}"]`);
    }
    form.querySelector('.inline-drawer-toggle')?.click();

    const changeList = [];
    const mapping = { 'selectiveLogic': 'entryLogicType' };
    const noParent = ['content', 'entryStateSelector', 'entryKillSwitch'];

    const hookInput = (key, inputs) => {
      for (const input of inputs) {
        const isSelect2 = input.classList.contains('select2-hidden-accessible') || input.style.display === 'none';

        const getTargetElement = () => {
          if (isSelect2) {
            const next = input.nextElementSibling;
            if (next && next.classList.contains('select2-container')) {
              return next.querySelector('.select2-selection') || next;
            }
          }
          if (noParent.includes(key)) return input;
          return input.parentElement;
        };

        const targetElement = getTargetElement();

        const addChange = () => {
          if (!changeList.includes(key)) {
            changeList.push(key);
            if (targetElement) targetElement.style.setProperty('outline', '2px solid orange', 'important');
          }
        };

        const removeChange = () => {
          if (changeList.includes(key)) {
            changeList.splice(changeList.indexOf(key), 1);
            if (targetElement) targetElement.style.removeProperty('outline');
          }
        };

        if (key !== 'entryKillSwitch') {
          $(input).on('input change select2:select select2:unselect', addChange);
        }

        const clickTarget = isSelect2 ? targetElement : input;
        if (clickTarget) {
          clickTarget.addEventListener('click', (evt) => {
            if (evt.ctrlKey) {
              evt.preventDefault();
              evt.stopImmediatePropagation();
              if (changeList.includes(key)) {
                removeChange();
              } else {
                addChange();
              }
            } else if (key === 'entryKillSwitch') {
              addChange();
            }
          }, true);
        }
      }
    };

    // Hook all possible inputs
    for (const key of Object.keys(newWorldInfoEntryDefinition)) {
      let input = [...form.querySelectorAll(`[name="${mapping[key] ?? key}"]`)];
      let snakeKey;
      if (!input.length) {
        snakeKey = key.replace(/([A-Z])/g, (_, c) => `_${c.toLowerCase()}`);
        input = [...form.querySelectorAll(`[name="${snakeKey}"]`)];
      }
      if (input.length) {
        hookInput(key, input);
      }
    }

    // Manually hook special ones
    hookInput('entryStateSelector', [...form.querySelectorAll('[name="entryStateSelector"]')]);
    hookInput('entryKillSwitch', [...form.querySelectorAll('[name="entryKillSwitch"]')]);
    hookInput('key', [...form.querySelectorAll('[name="key"]')]);
    hookInput('keysecondary', [...form.querySelectorAll('[name="keysecondary"]')]);
    hookInput('characterFilter', [
      ...form.querySelectorAll('[name="characterFilter"]'),
      ...form.querySelectorAll('[name="character_exclusion"]'),
    ]);

    // Build the dialog UI
    const dom = document.createElement('div');
    dom.classList.add('wis-bulk-dlg-main');

    const head = document.createElement('h3');
    head.textContent = `${i18n('bulkEditDialogTitle')} ${name}`;
    dom.append(head);

    const hint = document.createElement('small');
    hint.textContent = i18n('bulkEditHint');
    dom.append(hint);

    const contentWrapper = document.createElement('div');
    contentWrapper.classList.add('wis-bulk-content-wrapper');
    dom.append(contentWrapper);

    // Left panel: entry selection
    const selectionPanel = document.createElement('div');
    selectionPanel.classList.add('wis-bulk-selection-panel');

    const selectionHeader = document.createElement('h4');
    selectionHeader.textContent = i18n('bulkEditSelectEntries');
    selectionPanel.append(selectionHeader);

    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = i18n('bulkEditSearchPlaceholder');
    searchInput.classList.add('text_pole');
    searchInput.addEventListener('input', () => {
      const searchTerm = searchInput.value.toLowerCase();
      const labels = selectionPanel.querySelectorAll('.wis-bulk-entry-label');
      labels.forEach((label) => {
        const entryName = label.textContent.toLowerCase();
        label.style.display = entryName.includes(searchTerm) ? '' : 'none';
      });
    });
    selectionPanel.append(searchInput);

    const selectionActions = document.createElement('div');
    selectionActions.classList.add('wis-bulk-selection-actions');

    const selectAllBtn = document.createElement('div');
    selectAllBtn.textContent = i18n('bulkEditSelectAll');
    selectAllBtn.classList.add('menu_button', 'menu_button_small');
    selectAllBtn.addEventListener('click', () => {
      selectionPanel.querySelectorAll('.wis-bulk-entry-checkbox').forEach((cb) => cb.checked = true);
    });
    selectionActions.append(selectAllBtn);

    const deselectAllBtn = document.createElement('div');
    deselectAllBtn.textContent = i18n('bulkEditDeselectAll');
    deselectAllBtn.classList.add('menu_button', 'menu_button_small');
    deselectAllBtn.addEventListener('click', () => {
      selectionPanel.querySelectorAll('.wis-bulk-entry-checkbox').forEach((cb) => cb.checked = false);
    });
    selectionActions.append(deselectAllBtn);
    selectionPanel.append(selectionActions);

    const entryListContainer = document.createElement('div');
    entryListContainer.classList.add('wis-bulk-entry-list');

    Object.values(data.entries).filter((it) => it.uid !== entry.uid).forEach((e) => {
      const label = document.createElement('label');
      label.classList.add('wis-bulk-entry-label');

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.classList.add('wis-bulk-entry-checkbox');
      checkbox.value = e.uid;

      const text = document.createElement('span');
      text.textContent = `[${e.uid}] ${e.comment || e.key.join(', ')}`;
      text.title = `[${e.uid}] ${e.comment || e.key.join(', ')}`;

      label.append(checkbox, text);
      entryListContainer.append(label);
    });
    selectionPanel.append(entryListContainer);

    // Right panel: template
    const templatePanel = document.createElement('div');
    templatePanel.classList.add('wis-bulk-template-panel');

    const templateHeader = document.createElement('h4');
    templateHeader.textContent = i18n('bulkEditTemplate');
    templatePanel.append(templateHeader);
    form.querySelectorAll('.wis-entry-content-inline-copy-btn').forEach((btnElem) => btnElem.remove());
    templatePanel.append(form);

    contentWrapper.append(selectionPanel, templatePanel);

    let okToClose = false;
    let deleteTargets = false;
    let actionMode = 'apply';

    const dlg = new Popup(dom, POPUP_TYPE.CONFIRM, null, {
      okButton: i18n('bulkEditApply'),
      cancelButton: i18n('bulkEditCancel'),
      large: true,
      wider: true,
      allowVerticalScrolling: true,
      onClosing: () => okToClose,
      customButtons: [
        { text: i18n('bulkEditMoveCopy'), classes: ['wis-bulk-move-copy'] },
        { text: i18n('bulkEditCopyContents'), classes: ['wis-bulk-copy-contents'] },
        { text: i18n('bulkEditDelete'), classes: ['wis-bulk-delete', 'deleteworld_button'] },
      ],
    });

    const prom = dlg.show();

    // Handle OK button
    dlg.dlg.querySelector('.popup-button-ok').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((cb) => cb.value);

      if (selectedUids.length === 0) {
        toastr.warning(i18n('bulkEditNoEntriesSelected'));
        return;
      }
      if (changeList.length === 0 && !deleteTargets) {
        toastr.warning(i18n('bulkEditNoChanges'));
        return;
      }

      let confirmationHtml = `<h3>${i18n('bulkEditConfirmTitle')}</h3><p>${i18n('bulkEditConfirmMsg', selectedUids.length)}</p><ul>`;

      const tempBook = await loadWorldInfo(name);
      const templateEntry = tempBook.entries[entry.uid];

      for (const key of changeList) {
        let newValue = templateEntry[key];
        if (key === 'characterFilter') {
          if (newValue) {
            const names = newValue.names.join(', ') || i18n('labelNone');
            const tags = newValue.tags.join(', ') || i18n('labelNone');
            const mode = newValue.isExclude ? i18n('labelExclude') : i18n('labelOnly');
            newValue = `${i18n('labelMode')}: ${mode}, ${i18n('labelCharacters')}: ${names}, ${i18n('labelTags')}: ${tags}`;
          } else {
            newValue = `<i>${i18n('labelDisabled')}</i>`;
          }
          confirmationHtml += `<li><b>${key}</b>: <pre>${newValue}</pre></li>`;
          continue;
        }
        if (typeof newValue === 'boolean') {
          newValue = newValue ? i18n('labelYes') : i18n('labelNo');
        } else if (Array.isArray(newValue)) {
          newValue = newValue.join(', ');
        } else if (newValue === null || newValue === '') {
          newValue = `<i>${i18n('labelEmpty')}</i>`;
        }
        confirmationHtml += `<li><b>${key}</b>: <pre>${newValue}</pre></li>`;
      }
      confirmationHtml += `</ul><p>${i18n('bulkEditConfirmIrreversible')}</p>`;

      const confirmResult = await Popup.show.confirm(i18n('bulkEditConfirmTitle'), confirmationHtml, {
        okButton: i18n('bulkEditConfirmOk'),
        cancelButton: i18n('bulkEditConfirmCancel'),
      });

      if (confirmResult === POPUP_RESULT.AFFIRMATIVE) {
        okToClose = true;
        dlg.completeAffirmative();
      }
    });

    // Handle Move/Copy button
    dlg.dlg.querySelector('.wis-bulk-move-copy').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((cb) => cb.value);

      if (selectedUids.length === 0) {
        toastr.warning(i18n('bulkEditNoEntriesSelected'));
        return;
      }

      const moveCopyDom = document.createElement('div');
      const moveCopyHead = document.createElement('h3');
      moveCopyHead.textContent = i18n('moveCopyTitle');

      const moveCopyInfo = document.createElement('p');
      moveCopyInfo.textContent = i18n('moveCopyInfo', selectedUids.length);

      const targetWorldSelect = document.createElement('select');
      targetWorldSelect.classList.add('text_pole');
      targetWorldSelect.style.width = '100%';

      const sourceName = name;
      world_names.forEach((worldName) => {
        if (worldName !== sourceName) {
          const option = document.createElement('option');
          option.value = worldName;
          option.textContent = worldName;
          targetWorldSelect.append(option);
        }
      });

      if (targetWorldSelect.options.length === 0) {
        toastr.error(i18n('moveCopyNoTarget'));
        return;
      }

      moveCopyDom.append(moveCopyHead, moveCopyInfo, targetWorldSelect);

      const moveCopyPopup = new Popup(moveCopyDom, POPUP_TYPE.TEXT, null, {
        okButton: false,
        cancelButton: i18n('bulkEditCancel'),
        customButtons: [
          { text: i18n('moveCopyMove'), result: POPUP_RESULT.CUSTOM2 },
          { text: i18n('moveCopyCopy'), result: POPUP_RESULT.CUSTOM1 },
        ],
      });

      const moveCopyResult = await moveCopyPopup.show();

      if (!moveCopyResult) return;

      const targetName = targetWorldSelect.value;
      if (!targetName) {
        toastr.warning(i18n('moveCopyNoTargetSelected'));
        return;
      }

      const deleteOriginal = moveCopyResult === POPUP_RESULT.CUSTOM2;

      let successCount = 0;
      let errorCount = 0;
      for (const uid of selectedUids) {
        const success = await moveWorldInfoEntry(sourceName, targetName, uid, { deleteOriginal });
        if (success) {
          successCount++;
        } else {
          errorCount++;
        }
      }

      const action = deleteOriginal ? i18n('moveCopyActionMoved') : i18n('moveCopyActionCopied');
      toastr.success(i18n('moveCopySuccess', action, successCount, targetName));
      if (errorCount > 0) {
        toastr.error(i18n('moveCopyError', errorCount));
      }

      actionMode = 'moveCopy';
      okToClose = true;
      dlg.completeAffirmative();
    });

    // Handle Copy Entry Contents button
    dlg.dlg.querySelector('.wis-bulk-copy-contents').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((cb) => cb.value);
      if (selectedUids.length === 0) {
        toastr.warning(i18n('bulkCopyEntriesNoSelection'));
        return;
      }

      await copyEntryContentsByUid(data.entries, selectedUids);
    });

    // Handle Delete button
    dlg.dlg.querySelector('.wis-bulk-delete').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((cb) => cb.value);

      if (selectedUids.length === 0) {
        toastr.warning(i18n('bulkEditNoEntriesSelected'));
        return;
      }

      const confirmResult = await Popup.show.confirm(
        i18n('deleteConfirmTitle'),
        i18n('deleteConfirmMsg', selectedUids.length),
        { okButton: i18n('deleteConfirmOk'), cancelButton: i18n('bulkEditCancel') }
      );

      if (confirmResult) {
        actionMode = 'delete';
        okToClose = true;
        deleteTargets = true;
        dlg.completeAffirmative();
      }
    });

    // Handle Cancel button
    dlg.dlg.querySelector('.popup-button-cancel').addEventListener('click', () => {
      okToClose = true;
      dlg.completeNegative();
    });

    // Observer for Select2 containers
    const mo = new MutationObserver((muts) => {
      for (const mut of muts) {
        for (const n of [...mut.addedNodes].filter((it) => it instanceof HTMLElement && it.parentElement === document.body)) {
          if (!(n instanceof HTMLElement)) continue;
          if (n.classList.contains('select2-container')) {
            dlg.dlg.append(n);
          }
        }
      }
    });
    mo.observe(document.body, { childList: true, subtree: true });

    const outcome = await prom;
    mo.disconnect();

    if (outcome === POPUP_RESULT.AFFIRMATIVE) {
      const book = await loadWorldInfo(name);

      if (actionMode === 'moveCopy') {
        await deleteWorldInfoEntry(book, entry.uid, { silent: true });
        await saveWorldInfo(name, book, true);
      } else {
        const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((cb) => cb.value);
        const newEntry = book.entries[entry.uid];

        for (const uid of selectedUids) {
          if (deleteTargets) {
            await deleteWorldInfoEntry(book, uid, { silent: true });
          } else {
            const e = book.entries[uid];
            if (!e) continue;
            for (const key of changeList) {
              switch (key) {
                case 'entryStateSelector':
                  e.constant = newEntry.constant;
                  e.vectorized = newEntry.vectorized;
                  break;
                case 'entryKillSwitch':
                  e.disable = newEntry.disable;
                  break;
                case 'characterFilter':
                  if (newEntry.characterFilter) {
                    e.characterFilter = JSON.parse(JSON.stringify(newEntry.characterFilter));
                  } else {
                    delete e.characterFilter;
                  }
                  break;
                default:
                  if (Array.isArray(newEntry[key])) {
                    e[key] = [...newEntry[key]];
                  } else {
                    e[key] = newEntry[key];
                  }
                  break;
              }
            }
          }
        }

        await deleteWorldInfoEntry(book, entry.uid, { silent: true });
        await saveWorldInfo(name, book, true);

        const action = deleteTargets ? i18n('bulkEditActionDeleted') : i18n('bulkEditActionModified');
        toastr.success(i18n('bulkEditSuccess', action, selectedUids.length));
      }
    } else {
      // Cancelled - clean up the template entry
      const book = await loadWorldInfo(name);
      await deleteWorldInfoEntry(book, entry.uid, { silent: true });
      await saveWorldInfo(name, book, true);
    }

    sel.dispatchEvent(new Event('change', { bubbles: true }));
  });

  // Insert the button
  const anchor = document.querySelector('#world_apply_current_sorting');
  if (anchor) {
    anchor.insertAdjacentElement('afterend', btn);
  }

  // Update button visibility based on settings
  updateBulkEditorVisibility();
}

function updateBulkEditorVisibility() {
  const btn = document.getElementById('wis-bulk-edit-btn');
  if (btn) {
    btn.style.display = extension_settings.worldInfoSuite?.enableBulkEditor ? '' : 'none';
  }

  updateNativeEntryCopyButtonsVisibility();
}

// ============================================================
// Settings Panel
// ============================================================

async function loadSettings() {
  // Initialize extension settings if not present
  if (!extension_settings.worldInfoSuite) {
    extension_settings.worldInfoSuite = { ...defaultSettings };
  }

  // Merge with defaults for any missing keys
  extension_settings.worldInfoSuite = { ...defaultSettings, ...extension_settings.worldInfoSuite };
  getWorldbookManagerSettings();

  // Load and render settings HTML
  const settingsHtml = await renderExtensionTemplateAsync(extensionName, 'settings');
  $('#extensions_settings').append(settingsHtml);

  // Bind checkbox states
  $('#wis_enable_triggered_viewer').prop('checked', extension_settings.worldInfoSuite.enableTriggeredViewer);
  $('#wis_enable_char_lorebook').prop('checked', extension_settings.worldInfoSuite.enableCharLorebook);
  $('#wis_enable_bulk_editor').prop('checked', extension_settings.worldInfoSuite.enableBulkEditor);
  $('#wis_enable_worldbook_manager').prop('checked', extension_settings.worldInfoSuite.enableWorldbookManager);
  $('#wis_viewer_cache_limit').val(extension_settings.worldInfoSuite.viewerCacheLimit);

  // Bind change handlers
  $('#wis_enable_triggered_viewer').on('change', function () {
    extension_settings.worldInfoSuite.enableTriggeredViewer = $(this).prop('checked');
    saveSettingsDebounced();
  });

  $('#wis_enable_char_lorebook').on('change', function () {
    extension_settings.worldInfoSuite.enableCharLorebook = $(this).prop('checked');
    if (!$(this).prop('checked')) {
      hideCharacterWorldBooksPanel();
    }
    saveSettingsDebounced();
  });

  // Bind global lorebook toggles
  $('#wis_show_global_lorebook_mobile').prop('checked', extension_settings.worldInfoSuite.showGlobalLorebookMobile);
  $('#wis_show_global_lorebook_desktop').prop('checked', extension_settings.worldInfoSuite.showGlobalLorebookDesktop);

  $('#wis_show_global_lorebook_mobile').on('change', function () {
    extension_settings.worldInfoSuite.showGlobalLorebookMobile = $(this).prop('checked');
    saveSettingsDebounced();
  });

  $('#wis_show_global_lorebook_desktop').on('change', function () {
    extension_settings.worldInfoSuite.showGlobalLorebookDesktop = $(this).prop('checked');
    saveSettingsDebounced();
  });

  $('#wis_enable_bulk_editor').on('change', function () {
    extension_settings.worldInfoSuite.enableBulkEditor = $(this).prop('checked');
    updateBulkEditorVisibility();
    saveSettingsDebounced();
  });

  $('#wis_enable_worldbook_manager').on('change', function () {
    extension_settings.worldInfoSuite.enableWorldbookManager = $(this).prop('checked');
    updateWorldbookManagerVisibility();
    saveSettingsDebounced();
  });

  $('#wis_viewer_cache_limit').on('input', function () {
    const value = parseInt($(this).val(), 10);
    if (!isNaN(value) && value >= 0) {
      extension_settings.worldInfoSuite.viewerCacheLimit = value;
      saveSettingsDebounced();
      // Apply cleanup with new limit
      cleanupViewerCache();
    }
  });

  $('#wis_clear_cache_btn').on('click', function () {
    const count = clearAllViewerCache();
    if (count > 0) {
      toastr.success(i18n('cacheClearedSuccess', count));
    } else {
      toastr.info(i18n('cacheClearedNone'));
    }
  });

  // Bind viewer icon setting
  const currentIcon = extension_settings.worldInfoSuite.viewerIcon || 'fa-globe';
  $('#wis_viewer_icon').val(currentIcon);
  updateIconPreview(currentIcon);

  $('#wis_viewer_icon').on('change', function () {
    const icon = $(this).val();
    extension_settings.worldInfoSuite.viewerIcon = icon;
    updateIconPreview(icon);
    updateAllViewerButtonIcons(icon);
    saveSettingsDebounced();
  });
}

// Update icon preview in settings panel
function updateIconPreview(iconClass) {
  const preview = document.getElementById('wis_viewer_icon_preview');
  if (preview) {
    // Remove all fa-* classes except fa-regular
    preview.className = preview.className.replace(/fa-[\w-]+/g, '').trim();
    preview.classList.add('wis-icon-preview', 'fa-regular', iconClass);
  }
}

// Update all existing viewer buttons with new icon
function updateAllViewerButtonIcons(iconClass) {
  document.querySelectorAll('.worldinfo-viewer-btn').forEach((btn) => {
    // Remove old icon classes
    btn.className = btn.className.replace(/fa-[\w-]+/g, '').trim();
    btn.classList.add('mes_button', 'worldinfo-viewer-btn', 'fa-regular', iconClass);
  });
}

// ============================================================
// Initialization
// ============================================================

(async function init() {
  // Load localization first
  await loadLocaleData();

  // Load settings and render settings panel
  await loadSettings();

  // Initialize features
  initTriggeredViewer();
  initCharLorebookQuickAccess();
  initWorldbookManager();
  initBulkEditor();
  initNativeEntryCopyButtons();

  console.log(`[${extensionName}] World Info Suite initialized`);
})();
