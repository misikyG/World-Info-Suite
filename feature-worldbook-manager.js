import {
  eventSource,
  event_types,
  chat_metadata,
  characters,
  this_chid,
  saveSettingsDebounced,
} from '../../../../script.js';

import { extension_settings } from '../../../extensions.js';

import {
  METADATA_KEY,
  selected_world_info,
  world_info,
  world_names,
} from '../../../world-info.js';

import {
  Popup,
  POPUP_RESULT,
  POPUP_TYPE,
} from '../../../popup.js';

import { getCharaFilename } from '../../../utils.js';
import { defaultSettings, ensureSuiteSettings, i18n } from './core.js';

const WORLDBOOK_SORT_MODE = {
  CREATED_DESC: 'created_desc',
  CREATED_ASC: 'created_asc',
  NAME_ASC: 'name_asc',
  NAME_DESC: 'name_desc',
  CUSTOM: 'custom',
};

let worldbookManagerRefreshTimer = null;
let worldbookSelectObserver = null;
let worldbookUiObserver = null;
let worldbookPlaceholderOption;
let isApplyingWorldbookOptions = false;
let worldbookMutationSuppressedUntil = 0;

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

function normalizeKeywordToken(value) {
  return String(value ?? '').trim().toLocaleLowerCase();
}

function normalizeKeywordArray(value) {
  const arrayValue = Array.isArray(value)
    ? value
    : (typeof value === 'string' ? value.split(/[,\n]/) : []);

  const seen = new Set();
  const result = [];

  for (const raw of arrayValue) {
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
  if (!value || typeof value !== 'object') return output;

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
  ensureSuiteSettings();

  const suite = extension_settings.worldInfoSuite || { ...defaultSettings };
  const defaults = createWorldbookManagerDefaults();
  const raw = (suite.worldbookManager && typeof suite.worldbookManager === 'object')
    ? suite.worldbookManager
    : {};

  const normalized = {
    ...defaults,
    ...raw,
    customOrder: Array.isArray(raw.customOrder) ? raw.customOrder.filter((entry) => typeof entry === 'string') : [],
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
  if (normalized.activeTagFilter
    && !allTags.some((tag) => normalizeKeywordToken(tag) === normalizeKeywordToken(normalized.activeTagFilter))) {
    normalized.activeTagFilter = '';
  }

  suite.worldbookManager = normalized;
  extension_settings.worldInfoSuite = suite;
  return normalized;
}

function getWorldbookTags(worldName, settings = getWorldbookManagerSettings()) {
  if (!worldName) return [];
  return normalizeKeywordArray(settings?.worldTags?.[worldName] || []);
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
  const extraCharLore = world_info?.charLore?.find?.((entry) => entry.name === fileName);
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
    const aKeywordPriority = settings.priorityKeywords.some((keyword) => worldbookMatchesKeyword(a, keyword, settings)) ? 1 : 0;
    const bKeywordPriority = settings.priorityKeywords.some((keyword) => worldbookMatchesKeyword(b, keyword, settings)) ? 1 : 0;
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
  const select = document.querySelector('#world_editor_select');
  if (!(select instanceof HTMLSelectElement)) return '';
  if (select.value === '') return '';

  const index = Number(select.value);
  if (!Number.isNaN(index) && Array.isArray(world_names) && world_names[index]) {
    return world_names[index];
  }

  return '';
}

function rebuildWorldEditorOptions(orderedNames, selectedName) {
  const select = document.querySelector('#world_editor_select');
  if (!(select instanceof HTMLSelectElement) || !Array.isArray(world_names)) return;

  if (worldbookPlaceholderOption === undefined) {
    const sourcePlaceholder = select.querySelector('option[value=""]');
    worldbookPlaceholderOption = sourcePlaceholder ? sourcePlaceholder.cloneNode(true) : false;
    if (worldbookPlaceholderOption) {
      worldbookPlaceholderOption.value = '';
    }
  }

  const previousValue = select.value;
  const fragment = document.createDocumentFragment();
  if (worldbookPlaceholderOption) {
    fragment.append(worldbookPlaceholderOption.cloneNode(true));
  }

  const worldIndexMap = new Map(world_names.map((name, index) => [name, index]));
  const added = new Set();

  orderedNames.forEach((name) => {
    const index = worldIndexMap.get(name);
    if (index === undefined) return;

    const option = new Option(name, String(index));
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
  select.replaceChildren(fragment);
  isApplyingWorldbookOptions = false;

  if (selectedName && worldIndexMap.has(selectedName)) {
    const selectedValue = String(worldIndexMap.get(selectedName));
    if (select.querySelector(`option[value="${selectedValue}"]`)) {
      select.value = selectedValue;
    }
  }

  if (select.value === '' && select.querySelector('option[value=""]')) {
    select.value = '';
  }

  if (select.value !== previousValue) {
    select.dispatchEvent(new Event('change', { bubbles: true }));
  } else {
    $(select).trigger('change.select2');
  }
}

function applyWorldbookManagerToEditorSelect() {
  const select = document.querySelector('#world_editor_select');
  if (!(select instanceof HTMLSelectElement) || !Array.isArray(world_names) || world_names.length === 0) {
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
      const hiddenByKeyword = settings.hiddenKeywords.some((keyword) => worldbookMatchesKeyword(name, keyword, settings));
      if (hiddenByKeyword) return false;
    }

    if (hasActiveTagFilter && !worldbookHasTag(name, settings.activeTagFilter, settings)) {
      return false;
    }

    return true;
  });

  rebuildWorldEditorOptions(visibleNames, selectedName);
}

function ensureWorldbookVisibleForQuickJump(worldName) {
  if (!worldName || !Array.isArray(world_names) || world_names.length === 0) {
    return;
  }

  const select = document.querySelector('#world_editor_select');
  if (!(select instanceof HTMLSelectElement)) {
    return;
  }

  const worldIndex = world_names.indexOf(worldName);
  if (worldIndex < 0) {
    return;
  }

  const value = String(worldIndex);
  if (select.querySelector(`option[value="${value}"]`)) {
    return;
  }

  // Keep hidden lorebooks jumpable from external quick access entries.
  const label = extension_settings.worldInfoSuite?.enableWorldbookManager
    ? `${worldName} ${i18n('worldManagerHiddenSelectedSuffix')}`
    : worldName;

  select.append(new Option(label, value));
  $(select).trigger('change.select2');
}

function renderWorldbookQuickTagBar(select2Container = null) {
  const select = document.querySelector('#world_editor_select');
  if (!(select instanceof HTMLSelectElement)) return;

  const container = select2Container || select.nextElementSibling;
  if (!(container instanceof HTMLElement) || !container.classList.contains('select2')) return;

  const controlRow = container.closest('.flex-container.alignitemscenter');
  const host = container.parentElement || container;
  let tagBar = document.querySelector('.wis-worldbook-tag-bar');
  if (!(tagBar instanceof HTMLElement)) {
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
    chip.classList.add('menu_button', 'menu_button_small', 'interactable', 'wis-worldbook-tag-chip');
    if (isActive) {
      chip.classList.add('active');
    }

    chip.textContent = displayText;
    chip.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();

      const latest = getWorldbookManagerSettings();
      const isSameTag = normalizeKeywordToken(latest.activeTagFilter) === normalizeKeywordToken(tagValue);
      latest.activeTagFilter = isSameTag ? '' : tagValue;
      saveSettingsDebounced();
      applyWorldbookManagerToEditorSelect();
      renderWorldbookQuickTagBar(container);
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
  const select = document.querySelector('#world_editor_select');
  if (!(select instanceof HTMLSelectElement)) return false;

  const select2Container = select.nextElementSibling;
  if (!(select2Container instanceof HTMLElement) || !select2Container.classList.contains('select2')) {
    return false;
  }

  select2Container.classList.add('wis-worldbook-select2');

  const row = select2Container.parentElement;
  if (!(row instanceof HTMLElement)) {
    return false;
  }

  let manageButton = row.querySelector('.wis-worldbook-manage-btn');
  if (!(manageButton instanceof HTMLElement)) {
    manageButton = document.createElement('div');
    manageButton.classList.add('menu_button', 'fa-solid', 'fa-arrow-down-a-z', 'interactable', 'wis-worldbook-manage-btn');
    manageButton.setAttribute('role', 'button');
    manageButton.tabIndex = 0;

    const openManager = (event) => {
      event.preventDefault();
      event.stopPropagation();

      if (!extension_settings.worldInfoSuite?.enableWorldbookManager) {
        toastr.warning(i18n('worldManagerDisabled'));
        return;
      }

      showWorldbookManagerDialog();
    };

    manageButton.addEventListener('click', openManager);
    manageButton.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        openManager(event);
      }
    });
  }

  if (manageButton.parentElement !== row || manageButton.nextElementSibling !== select2Container) {
    row.insertBefore(manageButton, select2Container);
  }

  manageButton.title = i18n('worldManagerBtnTitle');
  manageButton.setAttribute('aria-label', i18n('worldManagerBtnTitle'));
  manageButton.style.display = extension_settings.worldInfoSuite?.enableWorldbookManager ? '' : 'none';

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

  let customSection;

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

  customSection = createSection('worldManagerSectionCustomOrder');
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
    const nameElement = document.createElement('span');
    nameElement.classList.add('wis-worldbook-order-name');
    nameElement.textContent = worldName;

    item.append(handle, nameElement);
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

    const nameElement = document.createElement('span');
    nameElement.classList.add('wis-worldbook-tag-editor-name');
    nameElement.textContent = worldName;

    const input = document.createElement('input');
    input.type = 'text';
    input.classList.add('text_pole');
    input.placeholder = i18n('worldManagerTagInputPlaceholder');
    input.value = getWorldbookTags(worldName, settings).join(', ');

    row.append(nameElement, input);
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
    } catch (_error) {
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
  if (settings.activeTagFilter
    && !availableTags.some((tag) => normalizeKeywordToken(tag) === normalizeKeywordToken(settings.activeTagFilter))) {
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
    const select = document.querySelector('#world_editor_select');
    if (!(select instanceof HTMLElement)) return;

    if (worldbookSelectObserver) {
      worldbookSelectObserver.disconnect();
    }

    worldbookSelectObserver = new MutationObserver(() => {
      if (isApplyingWorldbookOptions) return;
      if (Date.now() < worldbookMutationSuppressedUntil) return;
      scheduleWorldbookManagerRefresh();
    });

    worldbookSelectObserver.observe(select, { childList: true });
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
    const hasManageButton = row instanceof HTMLElement
      && Array.from(row.children).some((child) => child instanceof HTMLElement && child.classList.contains('wis-worldbook-manage-btn'));

    if (select2Container && !hasManageButton) {
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

  const manageButton = document.querySelector('.wis-worldbook-manage-btn');
  if (manageButton instanceof HTMLElement) {
    manageButton.style.display = enabled ? '' : 'none';
  }

  const tagBar = document.querySelector('.wis-worldbook-tag-bar');
  if (tagBar instanceof HTMLElement) {
    tagBar.style.display = enabled ? '' : 'none';
  }

  scheduleWorldbookManagerRefresh();
}

export {
  ensureWorldbookVisibleForQuickJump,
  getWorldbookManagerSettings,
  initWorldbookManager,
  updateWorldbookManagerVisibility,
};
