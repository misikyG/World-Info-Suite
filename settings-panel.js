import { saveSettingsDebounced } from '../../../../script.js';

import {
  extension_settings,
  renderExtensionTemplateAsync,
} from '../../../extensions.js';

import {
  extensionName,
  ensureSuiteSettings,
} from './core.js';

import {
  cleanupViewerCache,
  clearAllViewerCache,
  updateAllViewerButtonIcons,
} from './feature-triggered-viewer.js';

import {
  hideCharacterWorldBooksPanel,
} from './feature-char-lorebook.js';

import {
  updateBulkEditorVisibility,
} from './feature-bulk-editor.js';

import {
  updateEntryFoldersVisibility,
} from './feature-entry-folders.js';

import {
  getWorldbookManagerSettings,
  updateWorldbookManagerVisibility,
} from './feature-worldbook-manager.js';

import { i18n } from './core.js';

function updateIconPreview(iconClass) {
  const preview = document.getElementById('wis_viewer_icon_preview');
  if (preview) {
    preview.className = preview.className.replace(/fa-[\w-]+/g, '').trim();
    preview.classList.add('wis-icon-preview', 'fa-regular', iconClass);
  }
}

async function loadSettings() {
  ensureSuiteSettings();
  getWorldbookManagerSettings();

  const settingsHtml = await renderExtensionTemplateAsync(extensionName, 'settings');
  $('#extensions_settings').append(settingsHtml);

  $('#wis_enable_triggered_viewer').prop('checked', extension_settings.worldInfoSuite.enableTriggeredViewer);
  $('#wis_enable_char_lorebook').prop('checked', extension_settings.worldInfoSuite.enableCharLorebook);
  $('#wis_enable_bulk_editor').prop('checked', extension_settings.worldInfoSuite.enableBulkEditor);
  $('#wis_enable_entry_folders').prop('checked', extension_settings.worldInfoSuite.enableEntryFolders);
  $('#wis_enable_worldbook_manager').prop('checked', extension_settings.worldInfoSuite.enableWorldbookManager);
  $('#wis_viewer_cache_limit').val(extension_settings.worldInfoSuite.viewerCacheLimit);

  $('#wis_enable_triggered_viewer').on('change', function onTriggeredViewerToggle() {
    extension_settings.worldInfoSuite.enableTriggeredViewer = $(this).prop('checked');
    saveSettingsDebounced();
  });

  $('#wis_enable_char_lorebook').on('change', function onCharLorebookToggle() {
    extension_settings.worldInfoSuite.enableCharLorebook = $(this).prop('checked');
    if (!$(this).prop('checked')) {
      hideCharacterWorldBooksPanel();
    }
    saveSettingsDebounced();
  });

  $('#wis_show_global_lorebook_mobile').prop('checked', extension_settings.worldInfoSuite.showGlobalLorebookMobile);
  $('#wis_show_global_lorebook_desktop').prop('checked', extension_settings.worldInfoSuite.showGlobalLorebookDesktop);

  $('#wis_show_global_lorebook_mobile').on('change', function onGlobalMobileToggle() {
    extension_settings.worldInfoSuite.showGlobalLorebookMobile = $(this).prop('checked');
    saveSettingsDebounced();
  });

  $('#wis_show_global_lorebook_desktop').on('change', function onGlobalDesktopToggle() {
    extension_settings.worldInfoSuite.showGlobalLorebookDesktop = $(this).prop('checked');
    saveSettingsDebounced();
  });

  $('#wis_enable_bulk_editor').on('change', function onBulkEditorToggle() {
    extension_settings.worldInfoSuite.enableBulkEditor = $(this).prop('checked');
    updateBulkEditorVisibility();
    saveSettingsDebounced();
  });

  $('#wis_enable_entry_folders').on('change', function onEntryFoldersToggle() {
    extension_settings.worldInfoSuite.enableEntryFolders = $(this).prop('checked');
    updateEntryFoldersVisibility();
    saveSettingsDebounced();
  });

  $('#wis_enable_worldbook_manager').on('change', function onWorldbookManagerToggle() {
    extension_settings.worldInfoSuite.enableWorldbookManager = $(this).prop('checked');
    updateWorldbookManagerVisibility();
    saveSettingsDebounced();
  });

  $('#wis_viewer_cache_limit').on('input', function onViewerCacheLimitInput() {
    const value = parseInt($(this).val(), 10);
    if (!Number.isNaN(value) && value >= 0) {
      extension_settings.worldInfoSuite.viewerCacheLimit = value;
      saveSettingsDebounced();
      cleanupViewerCache();
    }
  });

  $('#wis_clear_cache_btn').on('click', function onClearCache() {
    const count = clearAllViewerCache();
    if (count > 0) {
      toastr.success(i18n('cacheClearedSuccess', count));
    } else {
      toastr.info(i18n('cacheClearedNone'));
    }
  });

  const currentIcon = extension_settings.worldInfoSuite.viewerIcon || 'fa-globe';
  $('#wis_viewer_icon').val(currentIcon);
  updateIconPreview(currentIcon);

  $('#wis_viewer_icon').on('change', function onViewerIconChange() {
    const icon = $(this).val();
    extension_settings.worldInfoSuite.viewerIcon = icon;
    updateIconPreview(icon);
    updateAllViewerButtonIcons(icon);
    saveSettingsDebounced();
  });
}

export {
  loadSettings,
};
