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
  viewerCacheLimit: 10, // Maximum number of messages to keep World Info viewer data
};

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

// ============================================================
// FEATURE 1: Triggered Entry Viewer
// ============================================================

function addViewButtonToMessage(messageId) {
  if (!extension_settings.worldInfoSuite?.enableTriggeredViewer) return;
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
  button.title = i18n('viewerBtnTitle');
  button.addEventListener('click', (event) => {
    event.stopPropagation();
    showWorldInfoPopup(messageId);
  });

  buttonContainer.prepend(button);
}

async function showWorldInfoPopup(messageId) {
  const worldInfoData = chat?.[messageId]?.extra?.worldInfoViewer;
  if (!worldInfoData) {
    toastr.info(i18n('noWorldInfoData'));
    return;
  }

  try {
    const popupContent = await renderExtensionTemplateAsync(extensionName, 'popup', {
      positions: worldInfoData,
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
    toastr.error('Unable to render World Info popup');
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
    const typeLabel = book.type === 'primary' ? i18n('charWorldbookTypePrimary') : i18n('charWorldbookTypeAdditional');
    const typeClass = book.type === 'primary' ? 'primary' : 'additional';
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
  eventSource.on(event_types.CHARACTER_EDITOR_OPENED, (chid) => {
    if (extension_settings.worldInfoSuite?.enableCharLorebook) {
      updateCharacterWorldBooksPanel(chid);
    }
  });

  eventSource.on(event_types.CHARACTER_EDITED, (data) => {
    const chid = data?.detail?.id;
    if (chid !== undefined && $('#char-worldbooks-panel').length) {
      if (extension_settings.worldInfoSuite?.enableCharLorebook) {
        updateCharacterWorldBooksPanel(chid);
      }
    }
  });
}

// ============================================================
// FEATURE 3: Bulk Entry Editor
// ============================================================

function initBulkEditor() {
  const btn = document.createElement('div');
  btn.id = 'wis-bulk-edit-btn';
  btn.classList.add('wis-bulk-trigger', 'menu_button', 'fa-solid', 'fa-list-check');
  btn.title = i18n('bulkEditBtnTitle');

  btn.addEventListener('click', async () => {
    if (!extension_settings.worldInfoSuite?.enableBulkEditor) {
      toastr.warning('Bulk Editor is disabled');
      return;
    }

    const sel = /** @type {HTMLSelectElement} */ (document.querySelector('#world_editor_select'));
    if (!sel || /** @type {HTMLOptionElement} */ (sel.children[0])?.selected) {
      toastr.warning(i18n('bulkEditNoWorldSelected'));
      return;
    }

    const name = sel.selectedOptions[0].textContent;
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
    templatePanel.append(form);

    contentWrapper.append(selectionPanel, templatePanel);

    let okToClose = false;
    let deleteTargets = false;

    const dlg = new Popup(dom, POPUP_TYPE.CONFIRM, null, {
      okButton: i18n('bulkEditApply'),
      cancelButton: i18n('bulkEditCancel'),
      large: true,
      wider: true,
      allowVerticalScrolling: true,
      onClosing: () => okToClose,
      customButtons: [
        { text: i18n('bulkEditMoveCopy'), classes: ['wis-bulk-move-copy'] },
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

      okToClose = true;
      dlg.completeAffirmative();
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
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((cb) => cb.value);
      const book = await loadWorldInfo(name);
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

  // Load and render settings HTML
  const settingsHtml = await renderExtensionTemplateAsync(extensionName, 'settings');
  $('#extensions_settings').append(settingsHtml);

  // Bind checkbox states
  $('#wis_enable_triggered_viewer').prop('checked', extension_settings.worldInfoSuite.enableTriggeredViewer);
  $('#wis_enable_char_lorebook').prop('checked', extension_settings.worldInfoSuite.enableCharLorebook);
  $('#wis_enable_bulk_editor').prop('checked', extension_settings.worldInfoSuite.enableBulkEditor);
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

  $('#wis_enable_bulk_editor').on('change', function () {
    extension_settings.worldInfoSuite.enableBulkEditor = $(this).prop('checked');
    updateBulkEditorVisibility();
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
  initBulkEditor();

  console.log(`[${extensionName}] World Info Suite initialized`);
})();
