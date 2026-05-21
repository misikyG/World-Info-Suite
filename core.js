import { extension_settings } from '../../../extensions.js';
import { addLocaleData, getCurrentLocale } from '../../../i18n.js';

const url = new URL(import.meta.url);
const extensionName = url.pathname.substring(url.pathname.lastIndexOf('extensions/') + 11, url.pathname.lastIndexOf('/'));
const extensionFolderPath = `scripts/extensions/${extensionName}`;

const defaultSettings = {
  enableTriggeredViewer: true,
  enableCharLorebook: true,
  enableBulkEditor: true,
  enableWorldbookManager: true,
  viewerCacheLimit: 10,
  viewerIcon: 'fa-globe',
  showGlobalLorebookMobile: true,
  showGlobalLorebookDesktop: true,
};

let localeData = {};

async function loadLocaleData() {
  const locale = getCurrentLocale();
  const localeToLoad = locale.startsWith('zh') ? 'zh-tw' : 'en';

  try {
    const response = await fetch(`${extensionFolderPath}/locales/${localeToLoad}.json`);
    if (response.ok) {
      localeData = await response.json();
      addLocaleData(locale, localeData);
      console.log(`[${extensionName}] Loaded locale: ${localeToLoad}`);
      return;
    }
  } catch (_error) {
    console.warn(`[${extensionName}] Failed to load locale ${localeToLoad}, falling back to en`);
  }

  try {
    const fallbackResponse = await fetch(`${extensionFolderPath}/locales/en.json`);
    if (fallbackResponse.ok) {
      localeData = await fallbackResponse.json();
      addLocaleData('en', localeData);
    }
  } catch (_error) {
    console.error(`[${extensionName}] Failed to load fallback locale`);
  }
}

function i18n(key, ...args) {
  let text = localeData[key] || key;
  args.forEach((arg, index) => {
    text = text.replace(new RegExp(`\\{${index}\\}`, 'g'), arg);
  });
  return text;
}

function getLocaleData() {
  return localeData;
}

function ensureSuiteSettings() {
  if (!extension_settings.worldInfoSuite) {
    extension_settings.worldInfoSuite = { ...defaultSettings };
  }

  extension_settings.worldInfoSuite = { ...defaultSettings, ...extension_settings.worldInfoSuite };
  return extension_settings.worldInfoSuite;
}

export {
  extensionName,
  extensionFolderPath,
  defaultSettings,
  loadLocaleData,
  i18n,
  getLocaleData,
  ensureSuiteSettings,
};
