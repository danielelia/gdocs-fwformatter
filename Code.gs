/**
 *  Pretty Code Add-On for Google Docs
 *
 *  adapted from Google's "Quickstart: Add-on for Google Docs"
 *  Tutorial at https://developers.google.com/apps-script/quickstart/docs
 *
 *  by Daniel Feist (April 2017)
 */

/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

 /**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Fixed Width Formatter');
  DocumentApp.getUi().showSidebar(ui);
}


// *
//  * Gets the stored user preferences for the origin and destination languages,
//  * if they exist.
//  * This method is only used by the regular add-on, and is never called by
//  * the mobile add-on version.
//  *
//  * @return {Object} The user's origin and destination language preferences, if
//  *     they exist.
 
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  var fontPrefs = {
    fixedWidthFont: userProperties.getProperty('fixedWidthFont'),
    fixedWidthFontSize: userProperties.getProperty('fixedWidthFontSize')
  };
  return fontPrefs;
}

/**
 * Replaces the text of the current selection with the provided text, or
 * inserts text at the current cursor location. (There will always be either
 * a selection or a cursor.) If multiple elements are selected, only inserts the
 * translated text in the first element that can contain text and removes the
 * other elements.
 *
 * @param {string} newText The text with which to replace the current selection.
 * @param {string} fixedWidthFont The font in question
 * @param {int} fontSize
 */
function formatText(fixedWidthFont, fixedWidthFontSize, savePrefs) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var fontSize = parseInt(fixedWidthFontSize);
  
  if (savePrefs == true) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('fixedWidthFont', fixedWidthFont);
    userProperties.setProperty('fixedWidthFontSize', fixedWidthFontSize);
  }
  
  if (selection) {
    var replaced = false;
    var elements = selection.getRangeElements();
    
    if (elements.length == 1 &&
        elements[0].getElement().getType() ==
        DocumentApp.ElementType.INLINE_IMAGE) {
      throw "Can't insert text into an image.";
    }

    for (var i = 0; i < elements.length; i++) {
      
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        element.setFontFamily(startIndex, endIndex, fixedWidthFont)
        element.setFontSize(startIndex, endIndex, fontSize);
        
      } else {
        
        var element = elements[i].getElement().asText();
        
        element.setFontFamily(fixedWidthFont)
        element.setFontSize(fontSize);
      }
      
    }
        
  }
}