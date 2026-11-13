import {
  eventSource,
  event_types,
  chat,
  chat_metadata,
  characters,
  this_chid,
  saveChatDebounced,
} from '../../../../script.js';

import {
  extension_settings,
  renderExtensionTemplateAsync,
} from '../../../extensions.js';

import {
  callGenericPopup,
  POPUP_TYPE,
} from '../../../popup.js';

import {
  METADATA_KEY,
  selected_world_info,
  world_info,
} from '../../../world-info.js';

import { getCharaFilename } from '../../../utils.js';
import { extensionName, i18n, getLocaleData } from './core.js';

const POSITION_SORT_ORDER = {
  0: 0,
  1: 1,
  5: 2,
  6: 3,
  2: 4,
  3: 5,
  4: 6,
  7: 7,
};

const WI_SOURCE_KEYS = {
  GLOBAL: 'global',
  CHARACTER_PRIMARY: 'characterPrimary',
  CHARACTER_ADDITIONAL: 'characterAdditional',
  CHAT: 'chat',
};

const ENTRY_SOURCE_TYPE = {
  ASSISTANT: 3,
  USER: 2,
  SYSTEM: 1,
};

const worldOrderCache = new Map();
let lastActivatedWorldInfo = null;

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

function getSelectiveLogicInfo() {
  return {
    0: i18n('selectiveLogicAndAny'),
    1: i18n('selectiveLogicNotAll'),
    2: i18n('selectiveLogicNotAny'),
    3: i18n('selectiveLogicAndAll'),
  };
}

function getWiSourceDisplay() {
  return {
    [WI_SOURCE_KEYS.GLOBAL]: i18n('sourceGlobal'),
    [WI_SOURCE_KEYS.CHARACTER_PRIMARY]: i18n('sourceCharacterPrimary'),
    [WI_SOURCE_KEYS.CHARACTER_ADDITIONAL]: i18n('sourceCharacterAdditional'),
    [WI_SOURCE_KEYS.CHAT]: i18n('sourceChat'),
  };
}

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
        const world = list[worldName];
        if (world && typeof world.order === 'number') {
          order = world.order;
          break;
        }
      }
    }
  } catch (_err) {
    // ignore
  }

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

  const sourceTypeA = a.sourceType ?? ENTRY_SOURCE_TYPE.ASSISTANT;
  const sourceTypeB = b.sourceType ?? ENTRY_SOURCE_TYPE.ASSISTANT;
  if (sourceTypeA !== sourceTypeB) return sourceTypeB - sourceTypeA;

  const orderA = (typeof a.worldOrder === 'number') ? a.worldOrder : Number.MAX_SAFE_INTEGER;
  const orderB = (typeof b.worldOrder === 'number') ? b.worldOrder : Number.MAX_SAFE_INTEGER;
  if (orderA !== orderB) return orderA - orderB;

  return String(a.entryName || '').localeCompare(String(b.entryName || ''));
}

function compareOrderEntries(a, b) {
  const orderA = (typeof a.worldOrder === 'number') ? a.worldOrder : Number.MAX_SAFE_INTEGER;
  const orderB = (typeof b.worldOrder === 'number') ? b.worldOrder : Number.MAX_SAFE_INTEGER;
  if (orderA !== orderB) return orderA - orderB;

  const worldNameCompare = String(a.worldName || '').localeCompare(String(b.worldName || ''));
  if (worldNameCompare !== 0) return worldNameCompare;
  return String(a.entryName || '').localeCompare(String(b.entryName || ''));
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
      (typeof entryRaw.worldOrder === 'number' ? entryRaw.worldOrder : undefined)
      ?? (typeof entryRaw.order === 'number' ? entryRaw.order : undefined)
      ?? getWorldOrderByName(entryRaw.world);

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

  Object.values(byPosition).forEach((positionGroup) => {
    if (positionGroup.position === 4) {
      positionGroup.entries.sort(compareDepthEntries);
    } else {
      positionGroup.entries.sort(compareOrderEntries);
    }
  });

  const groups = Object.values(byPosition).filter((group) => group.entries.length > 0);
  groups.sort((a, b) => getPositionSortIndex(a.position) - getPositionSortIndex(b.position));

  return groups;
}

function retranslateWorldInfoData(worldInfoData) {
  const positionInfo = getPositionInfo();
  const selectiveLogicInfo = getSelectiveLogicInfo();
  const statusEmojiMap = {
    '🔵': () => i18n('statusConstant'),
    '🔗': () => i18n('statusVectorized'),
    '🟢': () => i18n('statusKeyword'),
  };

  return worldInfoData.map((group) => {
    const posInfo = positionInfo[group.position] || { name: `${i18n('positionUnknown')} (${group.position})`, emoji: '❓' };
    return {
      ...group,
      positionName: posInfo.name,
      positionEmoji: posInfo.emoji,
      entries: group.entries.map((entry) => {
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

function addViewButtonToMessage(messageId) {
  if (!extension_settings.worldInfoSuite?.enableTriggeredViewer) return;
  if (!chat?.[messageId]?.extra?.worldInfoViewer) return;

  const messageElement = document.querySelector(`.mes[mesid="${messageId}"]`);
  if (!messageElement || messageElement.getAttribute('is_user') === 'true') return;

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

  extraButtonsContainer.prepend(button);
}

function setupViewerInteractions(root) {
  const groups = [...root.querySelectorAll('.wi-position-group')];
  const dotsHost = root.querySelector('.wi-nav-dots');
  const nav = root.querySelector('.wi-nav');
  if (!groups.length || !(dotsHost instanceof HTMLElement) || !(nav instanceof HTMLElement)) return;

  const toggleGroup = (group) => {
    const collapsed = group.classList.toggle('collapsed');
    group.querySelector('.wi-position-header')?.setAttribute('aria-expanded', String(!collapsed));
  };

  root.addEventListener('click', (event) => {
    const header = event.target instanceof Element ? event.target.closest('.wi-position-header') : null;
    if (header) toggleGroup(header.closest('.wi-position-group'));
  });

  root.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter' && event.key !== ' ') return;
    const header = event.target instanceof Element ? event.target.closest('.wi-position-header') : null;
    if (!header || event.target !== header) return;
    event.preventDefault();
    toggleGroup(header.closest('.wi-position-group'));
  });

  // A single section needs no navigation.
  if (groups.length < 2) {
    nav.style.display = 'none';
    return;
  }

  const dots = groups.map((group) => {
    const label = [
      group.querySelector('.wi-emoji')?.textContent,
      group.querySelector('.wi-position-name')?.textContent,
    ].map((text) => (text || '').trim()).filter(Boolean).join(' ');
    const dot = document.createElement('button');
    dot.type = 'button';
    dot.classList.add('wi-nav-dot');
    dot.title = label;
    dot.setAttribute('aria-label', label);
    dot.addEventListener('click', () => {
      group.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
    dotsHost.append(dot);
    return dot;
  });

  const setActive = (index) => dots.forEach((dot, i) => dot.classList.toggle('active', i === index));
  setActive(0);

  // The scroll container only exists once the popup has been attached to the document.
  let attempts = 0;
  const bindScroll = () => {
    const scroller = root.closest('.popup-content');
    if (!(scroller instanceof HTMLElement)) {
      attempts += 1;
      if (attempts < 40) setTimeout(bindScroll, 50);
      return;
    }

    const update = () => {
      const top = scroller.getBoundingClientRect().top;
      const atBottom = scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 2;
      let active = 0;
      groups.forEach((group, index) => {
        if (group.getBoundingClientRect().top - top <= 40) active = index;
      });
      setActive(atBottom && scroller.scrollTop > 0 ? groups.length - 1 : active);
    };

    scroller.addEventListener('scroll', update, { passive: true });
    update();
  };
  bindScroll();
}

async function showWorldInfoPopup(messageId) {
  const worldInfoData = chat?.[messageId]?.extra?.worldInfoViewer;
  if (!worldInfoData) {
    toastr.info(i18n('noWorldInfoData'));
    return;
  }

  const translatedData = retranslateWorldInfoData(worldInfoData);

  try {
    const popupHtml = await renderExtensionTemplateAsync(extensionName, 'popup', {
      positions: translatedData,
      i18n: getLocaleData(),
    });

    const popupContent = document.createElement('div');
    popupContent.innerHTML = popupHtml;
    setupViewerInteractions(popupContent);

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

function cleanupViewerCache() {
  if (!chat || !Array.isArray(chat)) return;

  const limit = extension_settings.worldInfoSuite?.viewerCacheLimit ?? 10;
  if (limit <= 0) return;

  const messagesWithData = [];
  for (let index = 0; index < chat.length; index += 1) {
    if (chat[index]?.extra?.worldInfoViewer) {
      messagesWithData.push(index);
    }
  }

  if (messagesWithData.length > limit) {
    const toRemove = messagesWithData.slice(0, messagesWithData.length - limit);
    for (const index of toRemove) {
      if (chat[index]?.extra?.worldInfoViewer) {
        delete chat[index].extra.worldInfoViewer;
      }
    }
    saveChatDebounced();
  }
}

function clearAllViewerCache() {
  if (!chat || !Array.isArray(chat)) return 0;

  let count = 0;
  for (let index = 0; index < chat.length; index += 1) {
    if (chat[index]?.extra?.worldInfoViewer) {
      delete chat[index].extra.worldInfoViewer;
      count += 1;
    }
  }

  document.querySelectorAll('.worldinfo-viewer-btn').forEach((button) => button.remove());

  if (count > 0) {
    saveChatDebounced();
  }

  return count;
}

function updateAllViewerButtonIcons(iconClass) {
  document.querySelectorAll('.worldinfo-viewer-btn').forEach((button) => {
    button.className = button.className.replace(/fa-[\w-]+/g, '').trim();
    button.classList.add('mes_button', 'worldinfo-viewer-btn', 'fa-regular', iconClass);
  });
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
      cleanupViewerCache();
    }
  });

  eventSource.on(event_types.CHARACTER_MESSAGE_RENDERED, (messageId) => {
    addViewButtonToMessage(String(messageId));
  });

  eventSource.on(event_types.CHAT_CHANGED, () => {
    setTimeout(() => {
      document.querySelectorAll('#chat .mes').forEach((messageElement) => {
        const messageId = messageElement.getAttribute('mesid');
        if (messageId) addViewButtonToMessage(messageId);
      });
    }, 500);
  });
}

export {
  initTriggeredViewer,
  cleanupViewerCache,
  clearAllViewerCache,
  updateAllViewerButtonIcons,
};
