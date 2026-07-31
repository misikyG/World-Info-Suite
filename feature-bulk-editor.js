import {
  eventSource,
  event_types,
} from '../../../../script.js';

import { extension_settings } from '../../../extensions.js';

import {
  POPUP_TYPE,
  Popup,
  POPUP_RESULT,
} from '../../../popup.js';

import {
  world_names,
  createWorldInfoEntry,
  deleteWorldInfoEntry,
  loadWorldInfo,
  saveWorldInfo,
  newWorldInfoEntryDefinition,
  moveWorldInfoEntry,
  sortWorldInfoEntries,
} from '../../../world-info.js';

import { delay } from '../../../utils.js';
import { remapEntryFolderUids } from './feature-entry-folders.js';
import { i18n } from './core.js';

let nativeEntryCopyObserver = null;
let nativeEntryCopyMountTimer = null;

const NATIVE_ENTRY_COPY_RETRY_DELAY = 45;
const NATIVE_ENTRY_COPY_MAX_RETRY = 12;
const BULK_TEMPLATE_FORM_RETRY_DELAY = 100;
const BULK_TEMPLATE_FORM_MAX_RETRY = 20;

async function copyTextToClipboard(text) {
  const normalized = String(text ?? '');

  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(normalized);
      return true;
    }
  } catch (_error) {
    // Fall through to execCommand fallback.
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
  } catch (_error) {
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

async function waitForBulkTemplateForm(uid) {
  for (let attempt = 0; attempt < BULK_TEMPLATE_FORM_MAX_RETRY; attempt += 1) {
    const form = document.querySelector(`#world_popup_entries_list .world_entry[uid="${uid}"]`);
    if (form instanceof HTMLElement) {
      return form;
    }

    await delay(BULK_TEMPLATE_FORM_RETRY_DELAY);
  }

  return null;
}

function formatBulkChangeValue(key, value) {
  if (key === 'characterFilter') {
    if (!value || typeof value !== 'object') {
      return i18n('labelDisabled');
    }

    const names = Array.isArray(value.names) ? value.names.join(', ') : '';
    const tags = Array.isArray(value.tags) ? value.tags.join(', ') : '';
    const mode = value.isExclude ? i18n('labelExclude') : i18n('labelOnly');
    return `${i18n('labelMode')}: ${mode}, ${i18n('labelCharacters')}: ${names || i18n('labelNone')}, ${i18n('labelTags')}: ${tags || i18n('labelNone')}`;
  }

  if (typeof value === 'boolean') {
    return value ? i18n('labelYes') : i18n('labelNo');
  }

  if (Array.isArray(value)) {
    return value.join(', ');
  }

  if (value === null || value === '') {
    return i18n('labelEmpty');
  }

  return String(value ?? '');
}

function createBulkEditConfirmation(changeList, templateEntry, selectedCount) {
  const content = document.createElement('div');
  const title = document.createElement('h3');
  title.textContent = i18n('bulkEditConfirmTitle');
  const description = document.createElement('p');
  description.textContent = i18n('bulkEditConfirmMsg', selectedCount);
  const changes = document.createElement('ul');

  changeList.forEach((key) => {
    const item = document.createElement('li');
    const label = document.createElement('b');
    label.textContent = `${key}: `;
    const value = document.createElement('pre');
    value.textContent = formatBulkChangeValue(key, templateEntry[key]);
    item.append(label, value);
    changes.append(item);
  });

  const warning = document.createElement('p');
  warning.textContent = i18n('bulkEditConfirmIrreversible');
  content.append(title, description, changes, warning);
  return content;
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

  let copyButton = leftGroup.querySelector('.wis-entry-content-inline-copy-btn');
  if (!(copyButton instanceof HTMLButtonElement)) {
    copyButton = document.createElement('button');
    copyButton.type = 'button';
    copyButton.classList.add('menu_button', 'menu_button_small', 'fa-regular', 'fa-copy', 'wis-entry-content-inline-copy-btn');
    copyButton.addEventListener('click', async (event) => {
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

  copyButton.title = i18n('bulkCopyEntriesSingleButtonTitle');
  copyButton.style.display = extension_settings.worldInfoSuite?.enableBulkEditor ? '' : 'none';

  const expandButton = leftGroup.querySelector('.editor_maximize');
  if (expandButton instanceof HTMLElement) {
    if (expandButton.nextElementSibling !== copyButton) {
      expandButton.insertAdjacentElement('afterend', copyButton);
    }
  } else if (copyButton.parentElement !== leftGroup) {
    leftGroup.append(copyButton);
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
  document.querySelectorAll('#world_popup_entries_list .wis-entry-content-inline-copy-btn').forEach((button) => {
    if (button instanceof HTMLElement) {
      button.style.display = enabled ? '' : 'none';
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
      const shouldMount = mutations.some((mutation) => mutation.type === 'childList'
        && (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0));
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
  if (!normalizedName) {
    return null;
  }

  return { select, name: normalizedName };
}

function initBulkEditor() {
  const reindexUidButton = document.createElement('div');
  reindexUidButton.id = 'wis-reindex-uid-btn';
  reindexUidButton.classList.add('wis-bulk-trigger', 'menu_button', 'fa-solid', 'fa-arrow-down-1-9');
  reindexUidButton.title = i18n('bulkReindexUidBtnTitle');

  reindexUidButton.addEventListener('click', async () => {
    if (!extension_settings.worldInfoSuite?.enableBulkEditor) {
      toastr.warning(i18n('bulkEditDisabled'));
      return;
    }

    const worldSelection = getSelectedWorldEditorState();
    if (!worldSelection) {
      toastr.warning(i18n('bulkEditNoWorldSelected'));
      return;
    }

    const { select, name } = worldSelection;
    const data = await loadWorldInfo(name);
    const entries = Object.values(data?.entries || {})
      .filter((item) => item && typeof item === 'object' && !Array.isArray(item))
      .map((item) => {
        const normalized = structuredClone(item);
        normalized.displayIndex = normalized.displayIndex ?? normalized.uid;
        return normalized;
      });

    if (entries.length === 0) {
      toastr.warning(i18n('bulkReindexUidNoEntries'));
      return;
    }

    const confirmResult = await Popup.show.confirm(
      i18n('bulkReindexUidConfirmTitle'),
      i18n('bulkReindexUidConfirmMsg', entries.length),
      {
        okButton: i18n('bulkReindexUidConfirmOk'),
        cancelButton: i18n('bulkEditCancel'),
      },
    );

    if (confirmResult !== POPUP_RESULT.AFFIRMATIVE) {
      return;
    }

    try {
      const sortedEntries = sortWorldInfoEntries(entries);
      const reorderedEntries = {};
      const uidMap = new Map();
      let changedCount = 0;

      sortedEntries.forEach((entryItem, index) => {
        const newUid = index;
        const newDisplayIndex = index;
        uidMap.set(String(entryItem.uid), String(newUid));
        if (Number(entryItem.uid) !== newUid || Number(entryItem.displayIndex) !== newDisplayIndex) {
          changedCount += 1;
        }

        entryItem.uid = newUid;
        entryItem.displayIndex = newDisplayIndex;
        reorderedEntries[newUid] = entryItem;
      });

      if (changedCount === 0) {
        toastr.info(i18n('bulkReindexUidNoChanges'));
        return;
      }

      data.entries = reorderedEntries;
      await saveWorldInfo(name, data, true);
  remapEntryFolderUids(name, uidMap);
      select.dispatchEvent(new Event('change', { bubbles: true }));
      toastr.success(i18n('bulkReindexUidSuccess', sortedEntries.length));
    } catch (error) {
      console.error('[WIS] Failed to reindex UIDs:', error);
      toastr.error(i18n('bulkReindexUidError'));
    }
  });

  const button = document.createElement('div');
  button.id = 'wis-bulk-edit-btn';
  button.classList.add('wis-bulk-trigger', 'menu_button', 'fa-solid', 'fa-list-check');
  button.title = i18n('bulkEditBtnTitle');

  button.addEventListener('click', async () => {
    if (!extension_settings.worldInfoSuite?.enableBulkEditor) {
      toastr.warning(i18n('bulkEditDisabled'));
      return;
    }

    const worldSelection = getSelectedWorldEditorState();
    if (!worldSelection) {
      toastr.warning(i18n('bulkEditNoWorldSelected'));
      return;
    }

    const { select, name } = worldSelection;
    const data = await loadWorldInfo(name);

    const entry = createWorldInfoEntry(null, data);
    if (!entry) {
      toastr.error(i18n('bulkEditTemplateLoadError'));
      return;
    }

    await saveWorldInfo(name, data, true);
    select.dispatchEvent(new Event('change', { bubbles: true }));

    const form = await waitForBulkTemplateForm(entry.uid);
    if (!form) {
      await deleteWorldInfoEntry(data, entry.uid, { silent: true });
      await saveWorldInfo(name, data, true);
      select.dispatchEvent(new Event('change', { bubbles: true }));
      toastr.error(i18n('bulkEditTemplateLoadError'));
      return;
    }

    form.querySelector('.inline-drawer-toggle')?.click();

    const changeList = [];
    const mapping = { selectiveLogic: 'entryLogicType' };
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
          clickTarget.addEventListener('click', (event) => {
            if (event.ctrlKey) {
              event.preventDefault();
              event.stopImmediatePropagation();
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

    for (const key of Object.keys(newWorldInfoEntryDefinition)) {
      let input = [...form.querySelectorAll(`[name="${mapping[key] ?? key}"]`)];
      if (!input.length) {
        const snakeKey = key.replace(/([A-Z])/g, (_, capture) => `_${capture.toLowerCase()}`);
        input = [...form.querySelectorAll(`[name="${snakeKey}"]`)];
      }
      if (input.length) {
        hookInput(key, input);
      }
    }

    hookInput('entryStateSelector', [...form.querySelectorAll('[name="entryStateSelector"]')]);
    hookInput('entryKillSwitch', [...form.querySelectorAll('[name="entryKillSwitch"]')]);
    hookInput('key', [...form.querySelectorAll('[name="key"]')]);
    hookInput('keysecondary', [...form.querySelectorAll('[name="keysecondary"]')]);
    hookInput('characterFilter', [
      ...form.querySelectorAll('[name="characterFilter"]'),
      ...form.querySelectorAll('[name="character_exclusion"]'),
    ]);

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

    const selectAllButton = document.createElement('div');
    selectAllButton.textContent = i18n('bulkEditSelectAll');
    selectAllButton.classList.add('menu_button', 'menu_button_small');
    selectAllButton.addEventListener('click', () => {
      selectionPanel.querySelectorAll('.wis-bulk-entry-checkbox').forEach((checkbox) => {
        checkbox.checked = true;
      });
    });
    selectionActions.append(selectAllButton);

    const deselectAllButton = document.createElement('div');
    deselectAllButton.textContent = i18n('bulkEditDeselectAll');
    deselectAllButton.classList.add('menu_button', 'menu_button_small');
    deselectAllButton.addEventListener('click', () => {
      selectionPanel.querySelectorAll('.wis-bulk-entry-checkbox').forEach((checkbox) => {
        checkbox.checked = false;
      });
    });
    selectionActions.append(deselectAllButton);
    selectionPanel.append(selectionActions);

    const entryListContainer = document.createElement('div');
    entryListContainer.classList.add('wis-bulk-entry-list');

    Object.values(data.entries).filter((item) => item.uid !== entry.uid).forEach((entryItem) => {
      const label = document.createElement('label');
      label.classList.add('wis-bulk-entry-label');

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.classList.add('wis-bulk-entry-checkbox');
      checkbox.value = entryItem.uid;

      const text = document.createElement('span');
      const entryLabel = `[${entryItem.uid}] ${getEntryDisplayTitle(entryItem)}`;
      text.textContent = entryLabel;
      text.title = entryLabel;

      label.append(checkbox, text);
      entryListContainer.append(label);
    });
    selectionPanel.append(entryListContainer);

    const templatePanel = document.createElement('div');
    templatePanel.classList.add('wis-bulk-template-panel');

    const templateHeader = document.createElement('h4');
    templateHeader.textContent = i18n('bulkEditTemplate');
    templatePanel.append(templateHeader);
    form.querySelectorAll('.wis-entry-content-inline-copy-btn').forEach((buttonElem) => buttonElem.remove());
    templatePanel.append(form);

    contentWrapper.append(selectionPanel, templatePanel);

    let okToClose = false;
    let deleteTargets = false;
    let actionMode = 'apply';

    const dialog = new Popup(dom, POPUP_TYPE.CONFIRM, null, {
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

    const promise = dialog.show();

    dialog.dlg.querySelector('.popup-button-ok').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((checkbox) => checkbox.value);

      if (selectedUids.length === 0) {
        toastr.warning(i18n('bulkEditNoEntriesSelected'));
        return;
      }
      if (changeList.length === 0 && !deleteTargets) {
        toastr.warning(i18n('bulkEditNoChanges'));
        return;
      }

      const tempBook = await loadWorldInfo(name);
      const templateEntry = tempBook?.entries?.[entry.uid];
      if (!templateEntry) {
        toastr.error(i18n('bulkEditTemplateLoadError'));
        return;
      }

      const confirmationPopup = new Popup(
        createBulkEditConfirmation(changeList, templateEntry, selectedUids.length),
        POPUP_TYPE.CONFIRM,
        null,
        {
        okButton: i18n('bulkEditConfirmOk'),
        cancelButton: i18n('bulkEditConfirmCancel'),
        },
      );
      const confirmResult = await confirmationPopup.show();

      if (confirmResult === POPUP_RESULT.AFFIRMATIVE) {
        okToClose = true;
        dialog.completeAffirmative();
      }
    });

    dialog.dlg.querySelector('.wis-bulk-move-copy').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((checkbox) => checkbox.value);

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
          successCount += 1;
        } else {
          errorCount += 1;
        }
      }

      const action = deleteOriginal ? i18n('moveCopyActionMoved') : i18n('moveCopyActionCopied');
      toastr.success(i18n('moveCopySuccess', action, successCount, targetName));
      if (errorCount > 0) {
        toastr.error(i18n('moveCopyError', errorCount));
      }

      actionMode = 'moveCopy';
      okToClose = true;
      dialog.completeAffirmative();
    });

    dialog.dlg.querySelector('.wis-bulk-copy-contents').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((checkbox) => checkbox.value);
      if (selectedUids.length === 0) {
        toastr.warning(i18n('bulkCopyEntriesNoSelection'));
        return;
      }

      await copyEntryContentsByUid(data.entries, selectedUids);
    });

    dialog.dlg.querySelector('.wis-bulk-delete').addEventListener('click', async () => {
      const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((checkbox) => checkbox.value);

      if (selectedUids.length === 0) {
        toastr.warning(i18n('bulkEditNoEntriesSelected'));
        return;
      }

      const confirmResult = await Popup.show.confirm(
        i18n('deleteConfirmTitle'),
        i18n('deleteConfirmMsg', selectedUids.length),
        { okButton: i18n('deleteConfirmOk'), cancelButton: i18n('bulkEditCancel') },
      );

      if (confirmResult) {
        actionMode = 'delete';
        okToClose = true;
        deleteTargets = true;
        dialog.completeAffirmative();
      }
    });

    dialog.dlg.querySelector('.popup-button-cancel').addEventListener('click', () => {
      okToClose = true;
      dialog.completeNegative();
    });

    const observer = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        for (const node of [...mutation.addedNodes].filter((entryNode) => entryNode instanceof HTMLElement && entryNode.parentElement === document.body)) {
          if (!(node instanceof HTMLElement)) continue;
          if (node.classList.contains('select2-container')) {
            dialog.dlg.append(node);
          }
        }
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });

    const outcome = await promise;
    observer.disconnect();

    if (outcome === POPUP_RESULT.AFFIRMATIVE) {
      const book = await loadWorldInfo(name);

      if (actionMode === 'moveCopy') {
        await deleteWorldInfoEntry(book, entry.uid, { silent: true });
        await saveWorldInfo(name, book, true);
      } else {
        const selectedUids = [...entryListContainer.querySelectorAll('.wis-bulk-entry-checkbox:checked')].map((checkbox) => checkbox.value);
        const newEntry = book.entries[entry.uid];

        for (const uid of selectedUids) {
          if (deleteTargets) {
            await deleteWorldInfoEntry(book, uid, { silent: true });
          } else {
            const worldEntry = book.entries[uid];
            if (!worldEntry) continue;

            for (const key of changeList) {
              switch (key) {
                case 'entryStateSelector':
                  worldEntry.constant = newEntry.constant;
                  worldEntry.vectorized = newEntry.vectorized;
                  break;
                case 'entryKillSwitch':
                  worldEntry.disable = newEntry.disable;
                  break;
                case 'characterFilter':
                  if (newEntry.characterFilter) {
                    worldEntry.characterFilter = JSON.parse(JSON.stringify(newEntry.characterFilter));
                  } else {
                    delete worldEntry.characterFilter;
                  }
                  break;
                default:
                  if (Array.isArray(newEntry[key])) {
                    worldEntry[key] = [...newEntry[key]];
                  } else {
                    worldEntry[key] = newEntry[key];
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
      const book = await loadWorldInfo(name);
      await deleteWorldInfoEntry(book, entry.uid, { silent: true });
      await saveWorldInfo(name, book, true);
    }

    select.dispatchEvent(new Event('change', { bubbles: true }));
  });

  const anchor = document.querySelector('#world_apply_current_sorting');
  if (anchor) {
    document.getElementById('wis-reindex-uid-btn')?.remove();
    document.getElementById('wis-bulk-edit-btn')?.remove();
    anchor.insertAdjacentElement('afterend', reindexUidButton);
    reindexUidButton.insertAdjacentElement('afterend', button);
  }

  updateBulkEditorVisibility();
}

function updateBulkEditorVisibility() {
  const reindexButton = document.getElementById('wis-reindex-uid-btn');
  if (reindexButton) {
    reindexButton.style.display = extension_settings.worldInfoSuite?.enableBulkEditor ? '' : 'none';
  }

  const bulkButton = document.getElementById('wis-bulk-edit-btn');
  if (bulkButton) {
    bulkButton.style.display = extension_settings.worldInfoSuite?.enableBulkEditor ? '' : 'none';
  }

  updateNativeEntryCopyButtonsVisibility();
}

export {
  initBulkEditor,
  updateBulkEditorVisibility,
  initNativeEntryCopyButtons,
};
