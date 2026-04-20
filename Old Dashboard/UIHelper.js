// ===== File: UIHelpers.gs (BEGIN) =====
// Headless-safe UI helpers.
// WHAT: Provide alert/prompt/toast wrappers that won’t crash in Web Apps or triggers.
// WHY: Web Apps run without a UI; these catch errors and log instead, so calls are safe.

function getUiSafe() {
  try {
    return SpreadsheetApp.getUi(); // works only in interactive contexts
  } catch (err) {
    Logger.log('[UI] getUi() not available in headless context: %s', err && err.message);
    return null;
  }
}

function alertSafe(message, title) {
  try {
    const ui = getUiSafe();
    if (ui) {
      ui.alert(title || 'Notice', String(message), ui.ButtonSet.OK);
    } else {
      Logger.log('[UI] ALERT: %s%s', title ? '[' + title + '] ' : '', String(message));
    }
  } catch (err) {
    Logger.log('[UI] alertSafe failed: %s', err && err.message);
  }
}

function promptSafe(message, defaultText) {
  const ui = getUiSafe();
  if (ui) {
    const res = ui.prompt(String(message), String(defaultText || ''), ui.ButtonSet.OK_CANCEL);
    return res && res.getSelectedButton() === ui.Button.OK ? res.getResponseText() : null;
  }
  Logger.log('[UI] PROMPT skipped (headless): %s', String(message));
  return defaultText || null;
}

function toastSafe(message, title, seconds) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(String(message), title || 'Info', Number(seconds || 5));
  } catch (err) {
    Logger.log('[UI] TOAST: %s%s', title ? '[' + title + '] ' : '', String(message));
  }
}
// ===== File: UIHelpers.gs (END) =====
