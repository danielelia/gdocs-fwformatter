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

/**
 * Gets the text the user has selected. If there is no selection,
 * this function displays an error message.
 *
 * @return {Array.<string>} The selected text.
 */
function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var text = [];
    var elements = selection.getSelectedElements();
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        var element = elements[i].getElement();
        // Only translate elements that can be edited as text; skip images and
        // other non-text elements.
        if (element.editAsText) {
          var elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText != '') {
            text.push(elementText);
          }
        }
      }
    }
    if (text.length == 0) {
      throw 'Please select some text.';
    }
    return text;
  } else {
    throw 'Please select some text.';
  }
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
    fontSize: userProperties.getProperty('fontSize')
  };
  return fontPrefs;
}

/**
 * Gets the user-selected text and changes the fonted to a specified fixed-
 * with one, as well as changing it to a small text size.
 *
 * @param {string} fixedWidthFont What fixed with font to convert the text into
 * @param {int} fontSize New size for the fixed width text
 * @param {boolean} savePrefs Whether to save the font size and fixed with font prefs
 */
function insertFormattedText(fixedWidthFont, fontSize, savePrefs) {  
  var result = {};
  var text = getSelectedText();
  result['text'] = text.join('\n');

  if (savePrefs == true) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('fontSize', fontSize);
    userProperties.setProperty('fixedWidthFont', fixedWidthFont);
  }

  insertText(text, fixedWidthFont, fontSize);
  
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
function insertText(newText, fixedWidthFont, fontSizeString) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var fontSize = parseInt(fontSizeString);
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