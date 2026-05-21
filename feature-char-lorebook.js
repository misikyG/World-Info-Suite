import {
  eventSource,
  event_types,
  chat_metadata,
  characters,
} from '../../../../script.js';

import { extension_settings } from '../../../extensions.js';

import {
  METADATA_KEY,
  selected_world_info,
  world_info,
  world_names,
  openWorldInfoEditor,
} from '../../../world-info.js';

import { getCharaFilename } from '../../../utils.js';
import { ensureWorldbookVisibleForQuickJump } from './feature-worldbook-manager.js';
import { i18n } from './core.js';

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

  const primaryWorld = character?.data?.extensions?.world;
  if (primaryWorld && world_names?.includes(primaryWorld)) {
    books.push({ name: primaryWorld, type: 'primary' });
  }

  const fileName = getCharaFilename?.(chid);
  const extraCharLore = world_info?.charLore?.find?.((entry) => entry.name === fileName);
  if (extraCharLore && Array.isArray(extraCharLore.extraBooks)) {
    extraCharLore.extraBooks.forEach((bookName) => {
      if (bookName && world_names?.includes(bookName)) {
        books.push({ name: bookName, type: 'additional' });
      }
    });
  }

  const chatLoreName = chat_metadata?.[METADATA_KEY];
  if (chatLoreName && world_names?.includes(chatLoreName)) {
    books.push({ name: chatLoreName, type: 'chat' });
  }

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
    let typeLabel;
    let typeClass;

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

  $('.char-worldbook-item').on('click', function onWorldbookClick() {
    const worldName = $(this).data('world-name');
    if (worldName && typeof openWorldInfoEditor === 'function') {
      ensureWorldbookVisibleForQuickJump(worldName);
      openWorldInfoEditor(worldName);
    }
  });
}

function hideCharacterWorldBooksPanel() {
  $('#char-worldbooks-panel').remove();
}

function initCharLorebookQuickAccess() {
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

  eventSource.on(event_types.WORLDINFO_SETTINGS_UPDATED, () => {
    if (currentEditorChid !== null && $('#char-worldbooks-panel').length) {
      if (extension_settings.worldInfoSuite?.enableCharLorebook) {
        updateCharacterWorldBooksPanel(currentEditorChid);
      }
    }
  });

  eventSource.on(event_types.CHAT_CHANGED, () => {
    if (currentEditorChid !== null && $('#char-worldbooks-panel').length) {
      if (extension_settings.worldInfoSuite?.enableCharLorebook) {
        updateCharacterWorldBooksPanel(currentEditorChid);
      }
    }
  });

  const popupObserver = new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      for (const node of mutation.removedNodes) {
        if (node instanceof HTMLElement && node.classList?.contains('popup')) {
          if (currentEditorChid !== null && $('#char-worldbooks-panel').length) {
            if (extension_settings.worldInfoSuite?.enableCharLorebook) {
              setTimeout(() => {
                updateCharacterWorldBooksPanel(currentEditorChid);
              }, 100);
            }
          }
        }
      }
    }
  });

  popupObserver.observe(document.body, { childList: true });
}

export {
  initCharLorebookQuickAccess,
  hideCharacterWorldBooksPanel,
};
