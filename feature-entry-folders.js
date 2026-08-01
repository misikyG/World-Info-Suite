import { saveSettingsDebounced } from '../../../../script.js';

import { extension_settings } from '../../../extensions.js';

import {
  POPUP_TYPE,
  Popup,
  POPUP_RESULT,
} from '../../../popup.js';

import {
  world_names,
  loadWorldInfo,
} from '../../../world-info.js';

import { i18n } from './core.js';

let entryListObserver = null;
let renderTimer = null;
let isRendering = false;

const DOM_MOUNT_RETRY_DELAY = 250;
const DOM_MOUNT_MAX_RETRY = 40;

function getSelectedWorldEditorState() {
  const select = document.querySelector('#world_editor_select');
  if (!(select instanceof HTMLSelectElement) || select.children[0]?.selected) {
    return null;
  }

  const selectedIndex = Number(select.value);
  const name = (!Number.isNaN(selectedIndex) && Array.isArray(world_names) && world_names[selectedIndex])
    ? world_names[selectedIndex]
    : select.selectedOptions[0]?.textContent;
  const normalizedName = String(name ?? '').trim();

  return normalizedName ? { name: normalizedName } : null;
}

function getEntryFoldersByWorld() {
  if (!extension_settings.worldInfoSuite.entryFolders || typeof extension_settings.worldInfoSuite.entryFolders !== 'object') {
    extension_settings.worldInfoSuite.entryFolders = {};
  }

  return extension_settings.worldInfoSuite.entryFolders;
}

function getFolders(worldName) {
  const foldersByWorld = getEntryFoldersByWorld();
  if (!Array.isArray(foldersByWorld[worldName])) {
    foldersByWorld[worldName] = [];
  }

  return foldersByWorld[worldName];
}

function saveFolders() {
  saveSettingsDebounced();
}

function remapEntryFolderUids(worldName, uidMap) {
  if (!worldName || !(uidMap instanceof Map) || uidMap.size === 0) {
    return false;
  }

  let changed = false;
  getFolders(worldName).forEach((folder) => {
    if (!Array.isArray(folder.uids)) {
      return;
    }

    const remappedUids = folder.uids.map((uid) => uidMap.get(String(uid)) ?? uid);
    if (remappedUids.some((uid, index) => uid !== folder.uids[index])) {
      folder.uids = remappedUids;
      changed = true;
    }
  });

  if (changed) {
    saveFolders();
  }

  return changed;
}

function getEntryUid(entry) {
  return String(entry.getAttribute('uid') ?? entry.getAttribute('data-uid') ?? '').trim();
}

function getEntryTitle(entry) {
  const comment = entry.querySelector('[name="comment"]');
  if (comment && 'value' in comment && String(comment.value).trim()) {
    return String(comment.value).trim();
  }

  return `${i18n('entryLabel')} #${getEntryUid(entry)}`;
}

function getDataEntryTitle(entry) {
  if (entry?.comment && String(entry.comment).trim()) {
    return String(entry.comment).trim();
  }

  if (Array.isArray(entry?.key) && entry.key.length > 0) {
    return entry.key.join(', ');
  }

  return `${i18n('entryLabel')} #${entry?.uid ?? ''}`;
}

function normalizeFolders(worldName, availableUids) {
  const seen = new Set();
  const folders = getFolders(worldName);
  let changed = false;

  folders.forEach((folder) => {
    const normalizedUids = [];
    (Array.isArray(folder.uids) ? folder.uids : []).forEach((uid) => {
      const normalizedUid = String(uid);
      if (!availableUids.has(normalizedUid) || seen.has(normalizedUid)) {
        return;
      }

      seen.add(normalizedUid);
      normalizedUids.push(normalizedUid);
    });

    const uidsChanged = !Array.isArray(folder.uids)
      || folder.uids.length !== normalizedUids.length
      || normalizedUids.some((uid, index) => uid !== folder.uids[index]);
    if (uidsChanged) {
      folder.uids = normalizedUids;
      changed = true;
    }

    const collapsed = folder.collapsed !== false;
    if (folder.collapsed !== collapsed) {
      folder.collapsed = collapsed;
      changed = true;
    }
  });

  return changed;
}

function assignEntryToFolder(worldName, folderId, uid) {
  const folders = getFolders(worldName);
  const targetFolder = folders.find((folder) => folder.id === folderId);
  if (!targetFolder) {
    return;
  }

  folders.forEach((folder) => {
    folder.uids = (folder.uids || []).map(String).filter((entryUid) => entryUid !== uid);
  });
  targetFolder.uids.push(uid);
  saveFolders();
  scheduleFolderRender();
}

function createTrashIcon() {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('wis-entry-folder-trash-icon');

  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', 'M4 7h16M10 11v6m4-6v6M9 7l1-3h4l1 3m-9 0 1 13h10l1-13');
  path.setAttribute('fill', 'none');
  path.setAttribute('stroke', 'currentColor');
  path.setAttribute('stroke-linecap', 'round');
  path.setAttribute('stroke-linejoin', 'round');
  path.setAttribute('stroke-width', '2');
  svg.append(path);
  return svg;
}

function createFolderIcon() {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('wis-entry-folder-icon');

  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', 'M3 7.5A2.5 2.5 0 0 1 5.5 5H10l2 2h6.5A2.5 2.5 0 0 1 21 9.5v8A2.5 2.5 0 0 1 18.5 20h-13A2.5 2.5 0 0 1 3 17.5z');
  path.setAttribute('fill', 'none');
  path.setAttribute('stroke', 'currentColor');
  path.setAttribute('stroke-linejoin', 'round');
  path.setAttribute('stroke-width', '2');
  svg.append(path);
  return svg;
}

function createDragHandleIcon() {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('wis-entry-folder-drag-handle');

  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', 'M5 7h14M5 12h14M5 17h14');
  path.setAttribute('fill', 'none');
  path.setAttribute('stroke', 'currentColor');
  path.setAttribute('stroke-linecap', 'round');
  path.setAttribute('stroke-width', '2');
  svg.append(path);
  return svg;
}

function createEditIcon() {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('aria-hidden', 'true');
  svg.classList.add('wis-entry-folder-edit-icon');

  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', 'm4 20 4.2-1 10.1-10.1a2.1 2.1 0 0 0-3-3L5.2 15.9zM13.8 7.4l3 3');
  path.setAttribute('fill', 'none');
  path.setAttribute('stroke', 'currentColor');
  path.setAttribute('stroke-linecap', 'round');
  path.setAttribute('stroke-linejoin', 'round');
  path.setAttribute('stroke-width', '2');
  svg.append(path);
  return svg;
}

async function showAddEntriesDialog(worldName, folder) {
  const data = await loadWorldInfo(worldName);
  const entries = Object.values(data?.entries || {})
    .filter((entry) => entry && typeof entry === 'object')
    .sort((first, second) => Number(first.displayIndex ?? first.uid) - Number(second.displayIndex ?? second.uid));

  if (entries.length === 0) {
    toastr.info(i18n('entryFolderNoEntries'));
    return;
  }

  const selectedUids = new Set((folder.uids || []).map(String));
  const content = document.createElement('div');
  content.classList.add('wis-entry-folder-picker');
  const heading = document.createElement('h3');
  heading.textContent = i18n('entryFolderAddEntriesTitle', folder.name);
  const list = document.createElement('div');
  list.classList.add('wis-entry-folder-picker-list');

  entries.forEach((entry) => {
    const uid = String(entry.uid);
    const label = document.createElement('label');
    label.classList.add('wis-entry-folder-picker-option');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = uid;
    checkbox.checked = selectedUids.has(uid);
    const title = document.createElement('span');
    title.textContent = getDataEntryTitle(entry);
    label.append(checkbox, title);
    list.append(label);
  });

  content.append(heading, list);
  const popup = new Popup(content, POPUP_TYPE.CONFIRM, null, {
    okButton: i18n('entryFolderAddEntriesConfirm'),
    cancelButton: i18n('bulkEditCancel'),
  });
  const result = await popup.show();
  if (result !== POPUP_RESULT.AFFIRMATIVE) {
    return;
  }

  const uids = [...list.querySelectorAll('input:checked')].map((checkbox) => checkbox.value);
  const folders = getFolders(worldName);
  folders.forEach((existingFolder) => {
    existingFolder.uids = (existingFolder.uids || []).map(String).filter((uid) => !uids.includes(uid));
  });
  folder.uids = uids;
  saveFolders();
  scheduleFolderRender();
}

function createFolderHeader(worldName, folder, entryCount) {
  const header = document.createElement('div');
  header.classList.add('wis-entry-folder');
  header.dataset.folderId = folder.id;
  header.draggable = true;

  const dragHandle = createDragHandleIcon();

  const toggle = document.createElement('button');
  toggle.type = 'button';
  toggle.classList.add(
    'wis-entry-folder-toggle',
    'fa-solid',
    folder.collapsed ? 'fa-circle-chevron-right' : 'fa-circle-chevron-down',
  );
  toggle.title = folder.collapsed ? i18n('entryFolderExpand') : i18n('entryFolderCollapse');

  const name = document.createElement('span');
  name.classList.add('wis-entry-folder-name');
  name.textContent = folder.name;

  const summary = document.createElement('span');
  summary.classList.add('wis-entry-folder-summary');
  summary.textContent = i18n('entryFolderSummary', entryCount);

  const add = document.createElement('button');
  add.type = 'button';
  add.classList.add('wis-entry-folder-action', 'menu_button', 'menu_button_small');
  add.title = i18n('entryFolderAddEntries');
  add.textContent = '+';

  const edit = document.createElement('button');
  edit.type = 'button';
  edit.classList.add('wis-entry-folder-action', 'wis-entry-folder-edit', 'menu_button', 'menu_button_small');
  edit.title = i18n('entryFolderRename');
  edit.append(createEditIcon());

  const remove = document.createElement('button');
  remove.type = 'button';
  remove.classList.add('wis-entry-folder-action', 'wis-entry-folder-delete', 'menu_button', 'menu_button_small');
  remove.title = i18n('entryFolderDelete');
  remove.append(createTrashIcon());

  toggle.addEventListener('click', (event) => {
    event.stopPropagation();
    folder.collapsed = !folder.collapsed;
    saveFolders();
    scheduleFolderRender();
  });
  add.addEventListener('click', async (event) => {
    event.stopPropagation();
    await showAddEntriesDialog(worldName, folder);
  });
  edit.addEventListener('click', async (event) => {
    event.stopPropagation();
    const name = (await Popup.show.input(
      i18n('entryFolderRenameTitle'),
      i18n('entryFolderRenamePrompt'),
      folder.name,
    ))?.trim();
    if (!name || name === folder.name) {
      return;
    }

    folder.name = name;
    saveFolders();
    scheduleFolderRender();
  });
  remove.addEventListener('click', async (event) => {
    event.stopPropagation();
    const confirmed = await Popup.show.confirm(
      i18n('entryFolderDeleteTitle'),
      i18n('entryFolderDeleteMessage', folder.name, entryCount),
      { okButton: i18n('entryFolderDelete'), cancelButton: i18n('bulkEditCancel') },
    );
    if (confirmed !== POPUP_RESULT.AFFIRMATIVE) {
      return;
    }

    const folders = getFolders(worldName);
    folders.splice(folders.findIndex((candidate) => candidate.id === folder.id), 1);
    saveFolders();
    scheduleFolderRender();
  });

  header.addEventListener('dragstart', (event) => {
    event.dataTransfer?.setData('application/x-wis-entry-folder', folder.id);
    event.dataTransfer?.setData('text/plain', folder.id);
    if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move';
  });
  header.addEventListener('dragover', (event) => {
    event.preventDefault();
    header.classList.add('wis-entry-folder-drag-over');
  });
  header.addEventListener('dragleave', () => header.classList.remove('wis-entry-folder-drag-over'));
  header.addEventListener('drop', (event) => {
    event.preventDefault();
    header.classList.remove('wis-entry-folder-drag-over');
    const entryUid = event.dataTransfer?.getData('application/x-wis-entry-uid');
    if (entryUid) {
      assignEntryToFolder(worldName, folder.id, entryUid);
      return;
    }

    const sourceFolderId = event.dataTransfer?.getData('application/x-wis-entry-folder');
    if (!sourceFolderId || sourceFolderId === folder.id) {
      return;
    }

    const folders = getFolders(worldName);
    const sourceIndex = folders.findIndex((candidate) => candidate.id === sourceFolderId);
    const targetIndex = folders.findIndex((candidate) => candidate.id === folder.id);
    if (sourceIndex < 0 || targetIndex < 0) {
      return;
    }

    const [sourceFolder] = folders.splice(sourceIndex, 1);
    folders.splice(targetIndex, 0, sourceFolder);
    saveFolders();
    scheduleFolderRender();
  });

  header.append(dragHandle, toggle, name, summary, add, edit, remove);
  return header;
}

function makeEntryDraggable(entry) {
  entry.draggable = false;
  const dragHandle = entry.querySelector('.drag-handle');
  if (!(dragHandle instanceof HTMLElement)) {
    return;
  }

  dragHandle.draggable = true;
  if (dragHandle.dataset.wisFolderDraggable === '1') {
    return;
  }

  dragHandle.dataset.wisFolderDraggable = '1';
  dragHandle.addEventListener('dragstart', (event) => {
    const uid = getEntryUid(entry);
    if (!uid || !event.dataTransfer) {
      return;
    }

    event.dataTransfer.setData('application/x-wis-entry-uid', uid);
    event.dataTransfer.setData('text/plain', uid);
    event.dataTransfer.effectAllowed = 'move';
  });
}

function renderFolders() {
  renderTimer = null;
  const entryList = document.querySelector('#world_popup_entries_list');
  const worldSelection = getSelectedWorldEditorState();
  if (!(entryList instanceof HTMLElement)) {
    return;
  }

  isRendering = true;
  entryListObserver?.disconnect();
  entryList.querySelectorAll('.wis-entry-folder').forEach((folder) => folder.remove());
  const entries = [...entryList.querySelectorAll(':scope > .world_entry')];
  entries.forEach((entry) => {
    entry.classList.remove('wis-entry-folder-member', 'wis-entry-folder-collapsed');
    entry.removeAttribute('data-wis-folder-id');
    entry.draggable = false;
    const dragHandle = entry.querySelector('.drag-handle');
    if (dragHandle instanceof HTMLElement) {
      dragHandle.draggable = false;
    }
  });

  if (!extension_settings.worldInfoSuite?.enableEntryFolders || !worldSelection) {
    isRendering = false;
    entryListObserver?.observe(entryList, { childList: true });
    return;
  }

  const entriesByUid = new Map(entries.map((entry) => [getEntryUid(entry), entry]).filter(([uid]) => uid));
  if (normalizeFolders(worldSelection.name, new Set(entriesByUid.keys()))) {
    saveFolders();
  }
  const folders = getFolders(worldSelection.name);
  const assignedUids = new Set();
  const fragment = document.createDocumentFragment();

  folders.forEach((folder) => {
    const folderEntries = (folder.uids || []).map(String)
      .map((uid) => entriesByUid.get(uid))
      .filter(Boolean);
    fragment.append(createFolderHeader(worldSelection.name, folder, folderEntries.length));
    folderEntries.forEach((entry) => {
      assignedUids.add(getEntryUid(entry));
      entry.classList.add('wis-entry-folder-member');
      entry.dataset.wisFolderId = folder.id;
      entry.classList.toggle('wis-entry-folder-collapsed', folder.collapsed);
      makeEntryDraggable(entry);
      fragment.append(entry);
    });
  });

  entries.filter((entry) => !assignedUids.has(getEntryUid(entry))).forEach((entry) => {
    makeEntryDraggable(entry);
    fragment.append(entry);
  });
  entryList.append(fragment);
  isRendering = false;
  entryListObserver?.observe(entryList, { childList: true });
}

function scheduleFolderRender() {
  if (renderTimer) {
    clearTimeout(renderTimer);
  }

  renderTimer = setTimeout(renderFolders, 20);
}

async function createFolder() {
  if (!extension_settings.worldInfoSuite?.enableEntryFolders) {
    toastr.warning(i18n('entryFolderDisabled'));
    return;
  }

  const worldSelection = getSelectedWorldEditorState();
  if (!worldSelection) {
    toastr.warning(i18n('bulkEditNoWorldSelected'));
    return;
  }

  const name = (await Popup.show.input(i18n('entryFolderCreateTitle'), i18n('entryFolderCreatePrompt')))?.trim();
  if (!name) {
    return;
  }

  getFolders(worldSelection.name).push({
    id: crypto.randomUUID?.() ?? `${Date.now()}-${Math.random()}`,
    name,
    uids: [],
    collapsed: true,
  });
  saveFolders();
  scheduleFolderRender();
}

function initEntryFolders() {
  const button = document.createElement('div');
  button.id = 'wis-entry-folder-create-btn';
  button.classList.add('menu_button', 'wis-entry-folder-create-btn');
  button.title = i18n('entryFolderCreateButton');
  button.setAttribute('aria-label', i18n('entryFolderCreateButton'));
  button.append(createFolderIcon());
  button.addEventListener('click', createFolder);

  const mount = () => {
    const bulkButton = document.getElementById('wis-bulk-edit-btn');
    if (bulkButton?.parentElement) {
      document.getElementById('wis-entry-folder-create-btn')?.remove();
      bulkButton.insertAdjacentElement('beforebegin', button);
      updateEntryFoldersVisibility();
      return true;
    }
    return false;
  };

  if (!mount()) {
    let retries = 0;
    const mountTimer = setInterval(() => {
      retries += 1;
      if (mount() || retries >= DOM_MOUNT_MAX_RETRY) clearInterval(mountTimer);
    }, DOM_MOUNT_RETRY_DELAY);
  }

  const bindObserver = () => {
    const entryList = document.querySelector('#world_popup_entries_list');
    if (!(entryList instanceof HTMLElement)) {
      return false;
    }

    entryListObserver?.disconnect();
    entryListObserver = new MutationObserver(() => {
      if (!isRendering) scheduleFolderRender();
    });
    entryListObserver.observe(entryList, { childList: true });
    document.querySelector('#world_editor_select')?.addEventListener('change', scheduleFolderRender);
    scheduleFolderRender();
    return true;
  };

  if (!bindObserver()) {
    let retries = 0;
    const observerTimer = setInterval(() => {
      retries += 1;
      if (bindObserver() || retries >= DOM_MOUNT_MAX_RETRY) clearInterval(observerTimer);
    }, DOM_MOUNT_RETRY_DELAY);
  }
}

function updateEntryFoldersVisibility() {
  const button = document.getElementById('wis-entry-folder-create-btn');
  if (button instanceof HTMLElement) {
    button.style.display = extension_settings.worldInfoSuite?.enableEntryFolders ? '' : 'none';
  }
  scheduleFolderRender();
}

export {
  initEntryFolders,
  updateEntryFoldersVisibility,
  remapEntryFolderUids,
};
