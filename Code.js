/**
 * Serves the main HTML application.
 */
function doGet(e) {
  // 1. Fetch the dynamic settings to get the App Name
  const props = PropertiesService.getScriptProperties();
  const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
  
  let dynamicTitle = 'Task Manager'; // Fallback title
  
  if (settingsStr) {
    try {
      const config = JSON.parse(settingsStr);
      if (config.appName) {
        dynamicTitle = config.appName;
      }
    } catch (err) {
      console.error("Failed to parse settings for title: ", err);
    }
  }

  // 2. Serve the HTML and set the TRUE browser tab title
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(dynamicTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // Standard GAS setting
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Safe include helper for modular HTML.
 * Prevents the need for a giant ClientScript.html.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}