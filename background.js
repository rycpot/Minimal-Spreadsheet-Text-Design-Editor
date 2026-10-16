// Opens the editor as a full browser tab (not a popup) when the toolbar icon is clicked.
// If an editor tab is already open, focus it instead of opening a duplicate.

chrome.action.onClicked.addListener(async () => {
  const editorUrl = chrome.runtime.getURL('editor.html');

  const tabs = await chrome.tabs.query({ url: editorUrl });
  if (tabs.length > 0) {
    const tab = tabs[0];
    await chrome.tabs.update(tab.id, { active: true });
    await chrome.windows.update(tab.windowId, { focused: true });
    return;
  }

  await chrome.tabs.create({ url: editorUrl });
});
