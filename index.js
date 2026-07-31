import {
  extensionName,
  loadLocaleData,
} from './core.js';

import {
  loadSettings,
} from './settings-panel.js';

import {
  initTriggeredViewer,
} from './feature-triggered-viewer.js';

import {
  initCharLorebookQuickAccess,
} from './feature-char-lorebook.js';

import {
  initWorldbookManager,
} from './feature-worldbook-manager.js';

import {
  initBulkEditor,
  initNativeEntryCopyButtons,
} from './feature-bulk-editor.js';

import {
  initEntryFolders,
} from './feature-entry-folders.js';

(async function init() {
  await loadLocaleData();
  await loadSettings();

  initTriggeredViewer();
  initCharLorebookQuickAccess();
  initWorldbookManager();
  initBulkEditor();
  initEntryFolders();
  initNativeEntryCopyButtons();

  console.log(`[${extensionName}] World Info Suite initialized`);
}());
