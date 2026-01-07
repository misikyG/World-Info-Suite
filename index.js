import {
  eventSource,
  event_types,
  chat,
  chat_metadata,
  characters,
  this_chid,
} from '../../../../script.js';

import {
  renderExtensionTemplateAsync
} from '../../../extensions.js';

import {
  callGenericPopup,
  POPUP_TYPE
} from '../../../popup.js';

import {
  METADATA_KEY,
  selected_world_info,
  world_info,
  world_names,
  openWorldInfoEditor,
} from '../../../world-info.js';

import { getCharaFilename } from '../../../utils.js';

const url = new URL(import.meta.url);
const extensionName = url.pathname.substring(url.pathname.lastIndexOf('extensions/') + 11, url.pathname.lastIndexOf('/'));

const positionInfo = {
  0: { name: '角色設定前', emoji: '📙' },
  1: { name: '角色設定後', emoji: '📙' },
  2: { name: '作者註釋前', emoji: '📝' },
  3: { name: '作者註釋後', emoji: '📝' },
  4: { name: '依深度插入', emoji: '💉' },
  5: { name: '範例前', emoji: '📄' },
  6: { name: '範例後', emoji: '📄' },
  7: { name: 'Outlet', emoji: '➡️' },
};

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

const selectiveLogicInfo = {
  0: '包含任一 (AND ANY)',
  1: '未完全包含 (NOT ALL)',
  2: '完全不含 (NOT ANY)',
  3: '包含全部 (AND ALL)',
};

const WI_SOURCE_KEYS = {
  GLOBAL: 'global',
  CHARACTER_PRIMARY: 'characterPrimary',
  CHARACTER_ADDITIONAL: 'characterAdditional',
  CHAT: 'chat',
};

const WI_SOURCE_DISPLAY = {
  [WI_SOURCE_KEYS.GLOBAL]: '全域',
  [WI_SOURCE_KEYS.CHARACTER_PRIMARY]: '主要知識',
  [WI_SOURCE_KEYS.CHARACTER_ADDITIONAL]: '額外知識',
  [WI_SOURCE_KEYS.CHAT]: '聊天知識',
};

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
  if (role === 'assistant') return '@AI';
  if (role === 'user') return '@使用者';
  if (role === 'system') return '@系統';
  return '@AI';
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
  return WI_SOURCE_DISPLAY[sourceKey] || '';
}

function getEntryStatus(entry) {
  if (entry.constant === true) return { emoji: '🔵', name: '恆定 (Constant)' };
  if (entry.vectorized === true) return { emoji: '🔗', name: '向量 (Vectorized)' };
  return { emoji: '🟢', name: '關鍵字 (Keyword)' };
}

function formatRoleDepthTag(entry) {
  const roleString = getRoleString(entry.role);
  const depth = entry.depth ?? null;
  if (depth == null) return '';
  return `${roleDisplayName(roleString)} 深度 ${depth}`;
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
  } catch (err) {
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

function processWorldInfoData(activatedEntries) {
  const byPosition = {};

  activatedEntries.forEach((entryRaw) => {
    if (!entryRaw || typeof entryRaw !== 'object') return;

    const position = (typeof entryRaw.position === 'number') ? entryRaw.position : 0;
    const posInfo = positionInfo[position] || { name: `未知位置 (${position})`, emoji: '❓' };
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
      entryName: entryRaw.comment || `條目 #${entryRaw.uid}`,
      sourceKey,
      sourceName,
      statusEmoji: status.emoji,
      statusName: status.name,
      content: entryRaw.content,
      keys: Array.isArray(entryRaw.key) ? entryRaw.key.join(', ') : (typeof entryRaw.key === 'string' ? entryRaw.key : null),
      secondaryKeys: Array.isArray(entryRaw.keysecondary) ? entryRaw.keysecondary.join(', ') : null,
      selectiveLogicName: Array.isArray(entryRaw.keysecondary)
        ? (selectiveLogicInfo?.[entryRaw.selectiveLogic] ?? `未知邏輯 (${entryRaw.selectiveLogic})`)
        : null,
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

// ===== UI：訊息按鈕 =====
function addViewButtonToMessage(messageId) {
  if (!chat?.[messageId]?.extra?.worldInfoViewer) return;

  const messageElement = document.querySelector(`.mes[mesid="${messageId}"]`);
  if (!messageElement || messageElement.getAttribute('is_user') === 'true') return;

  const buttonContainer = messageElement.querySelector('.mes_buttons');
  if (!buttonContainer) return;

  const buttonId = `worldinfo-viewer-btn-${messageId}`;
  if (document.getElementById(buttonId)) return;

  const button = document.createElement('div');
  button.id = buttonId;
  button.className = 'mes_button worldinfo-viewer-btn fa-regular fa-globe';
  button.title = '查看此回覆觸發的世界書';
  button.addEventListener('click', (event) => {
    event.stopPropagation();
    showWorldInfoPopup(messageId);
  });

  buttonContainer.prepend(button);
}

// ===== UI：彈窗 =====
async function showWorldInfoPopup(messageId) {
  const worldInfoData = chat?.[messageId]?.extra?.worldInfoViewer;
  if (!worldInfoData) {
    toastr.info('此訊息沒有紀錄的世界書觸發資料。');
    return;
  }

  try {
    const popupContent = await renderExtensionTemplateAsync(extensionName, 'popup', { positions: worldInfoData });
    callGenericPopup(popupContent, POPUP_TYPE.TEXT, '', {
      wide: true,
      large: true,
      okButton: '關閉',
      allowVerticalScrolling: true,
    });
  } catch (error) {
    console.error(`[${extensionName}] 渲染彈窗時發生錯誤:`, error);
    toastr.error('無法渲染世界書彈窗，請檢查主控台日誌。');
  }
}

// ===== 狀態同步 =====
let lastActivatedWorldInfo = null;

eventSource.on(event_types.WORLD_INFO_ACTIVATED, (data) => {
  if (data && Array.isArray(data) && data.length > 0) {
    lastActivatedWorldInfo = processWorldInfoData(data);
  } else {
    lastActivatedWorldInfo = null;
  }
});

eventSource.on(event_types.MESSAGE_RECEIVED, (messageId) => {
  if (lastActivatedWorldInfo && chat?.[messageId] && !chat[messageId].is_user) {
    if (!chat[messageId].extra) chat[messageId].extra = {};
    chat[messageId].extra.worldInfoViewer = lastActivatedWorldInfo;
    lastActivatedWorldInfo = null;
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

// ===== 角色世界書列表功能 =====

/**
 * 取得角色綁定的所有世界書
 * @param {number} chid 角色 ID
 * @returns {Array<{name: string, type: string}>} 世界書列表
 */
function getCharacterWorldBooks(chid) {
  const books = [];
  const character = characters?.[chid];
  
  if (!character) return books;
  
  // 主要世界書 (Primary Lorebook)
  const primaryWorld = character?.data?.extensions?.world;
  if (primaryWorld && world_names?.includes(primaryWorld)) {
    books.push({ name: primaryWorld, type: 'primary' });
  }
  
  // 額外世界書 (Additional Lorebooks)
  const fileName = getCharaFilename?.(chid);
  const extraCharLore = world_info?.charLore?.find?.((e) => e.name === fileName);
  if (extraCharLore && Array.isArray(extraCharLore.extraBooks)) {
    extraCharLore.extraBooks.forEach((bookName) => {
      if (bookName && world_names?.includes(bookName)) {
        books.push({ name: bookName, type: 'additional' });
      }
    });
  }
  
  return books;
}

/**
 * 建立角色世界書視窗的 HTML
 * @param {Array<{name: string, type: string}>} books 世界書列表
 * @returns {string} HTML 字串
 */
function createCharacterWorldBooksHTML(books) {
  if (books.length === 0) {
    return `
      <div id="char-worldbooks-panel" class="char-worldbooks-panel">
        <div class="char-worldbooks-header">
          <span class="fa-solid fa-globe"></span>
          <span>綁定的世界書</span>
        </div>
        <div class="char-worldbooks-empty">
          此角色未綁定任何世界書
        </div>
      </div>
    `;
  }
  
  const bookItems = books.map((book) => {
    const typeLabel = book.type === 'primary' ? '主要' : '額外';
    const typeClass = book.type === 'primary' ? 'primary' : 'additional';
    return `
      <div class="char-worldbook-item" data-world-name="${book.name}">
        <span class="char-worldbook-type ${typeClass}">${typeLabel}</span>
        <span class="char-worldbook-name">${book.name}</span>
        <span class="char-worldbook-goto fa-solid fa-arrow-up-right-from-square" title="跳轉到世界書"></span>
      </div>
    `;
  }).join('');
  
  return `
    <div id="char-worldbooks-panel" class="char-worldbooks-panel">
      <div class="char-worldbooks-header">
        <span class="fa-solid fa-globe"></span>
        <span>綁定的世界書</span>
      </div>
      <div class="char-worldbooks-list">
        ${bookItems}
      </div>
    </div>
  `;
}

/**
 * 更新角色世界書面板
 * @param {number} chid 角色 ID
 */
function updateCharacterWorldBooksPanel(chid) {
  // 移除舊的面板
  $('#char-worldbooks-panel').remove();
  
  // 取得角色的世界書
  const books = getCharacterWorldBooks(chid);
  
  // 建立新的面板
  const panelHTML = createCharacterWorldBooksHTML(books);
  
  // 找到 char-management-dropdown 的父元素並在其後插入
  const dropdownLabel = $('#char-management-dropdown').closest('label');
  if (dropdownLabel.length) {
    dropdownLabel.after(panelHTML);
  } else {
    // 備用：直接插入到 avatar_controls
    $('#avatar_controls').append(panelHTML);
  }
  
  // 綁定點擊事件
  $('.char-worldbook-item').on('click', function() {
    const worldName = $(this).data('world-name');
    if (worldName && typeof openWorldInfoEditor === 'function') {
      openWorldInfoEditor(worldName);
    }
  });
}

/**
 * 隱藏角色世界書面板
 */
function hideCharacterWorldBooksPanel() {
  $('#char-worldbooks-panel').remove();
}

// 監聽角色編輯器開啟事件
eventSource.on(event_types.CHARACTER_EDITOR_OPENED, (chid) => {
  updateCharacterWorldBooksPanel(chid);
});

// 監聽角色編輯事件（世界書可能被修改）
eventSource.on(event_types.CHARACTER_EDITED, (data) => {
  const chid = data?.detail?.id;
  if (chid !== undefined && $('#char-worldbooks-panel').length) {
    updateCharacterWorldBooksPanel(chid);
  }
});