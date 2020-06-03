function createSolutionTemplate() {
  
  // Define vars for spreadsheet name, sheet name and the target solution doc
  var spreadsheet_id = "1VvRUa00mWChyXxA2nzFddRZSzMuSJ8VShxeAH7VIWlw";
  var sheet_name = SpreadsheetApp.getActiveSheet().getSheetName();
  
  // Open spreadsheet and get last row for the sheet
  var spreadsheet = SpreadsheetApp.openById(spreadsheet_id);

  Logger.log(sheet_name);
  var sheet = spreadsheet.getSheetByName(sheet_name);
  var sheet_last_row = sheet.getLastRow();
  var sheet_data = sheet.getRange(10, 1, sheet_last_row-4, 4).getValues();
  
  // Define styles for code and text
  var code_style = {};
  code_style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  code_style[DocumentApp.Attribute.FONT_FAMILY] = 'Inconsolata';
  code_style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#3c78d8';
  code_style[DocumentApp.Attribute.FONT_SIZE] = 10;
  code_style[DocumentApp.Attribute.BOLD] = false;
  code_style[DocumentApp.Attribute.INDENT_FIRST_LINE] = 36;
  code_style[DocumentApp.Attribute.INDENT_START] = 36;

  var output_style = {};
  output_style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  output_style[DocumentApp.Attribute.FONT_FAMILY] = 'Inconsolata';
  output_style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#666666';
  output_style[DocumentApp.Attribute.FONT_SIZE] = 10;
  output_style[DocumentApp.Attribute.BOLD] = false;
  output_style[DocumentApp.Attribute.INDENT_FIRST_LINE] = 36;
  output_style[DocumentApp.Attribute.INDENT_START] = 36;
  
  var steps_style = {};
  steps_style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  steps_style[DocumentApp.Attribute.FONT_FAMILY] = 'Google Sans';
  steps_style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#434343';
  steps_style[DocumentApp.Attribute.FONT_SIZE] = 10;
  steps_style[DocumentApp.Attribute.BOLD] = false;
  steps_style[DocumentApp.Attribute.SPACING_AFTER] = 12;
  steps_style[DocumentApp.Attribute.SPACING_BEFORE] = 12;
  
  var author_style = {};
  author_style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  author_style[DocumentApp.Attribute.FONT_FAMILY] = 'Google Sans';
  author_style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#999999';
  author_style[DocumentApp.Attribute.FONT_SIZE] = 8;
  author_style[DocumentApp.Attribute.BOLD] = false;
  author_style[DocumentApp.Attribute.SPACING_AFTER] = 12;
  author_style[DocumentApp.Attribute.SPACING_BEFORE] = 12;
  
  var note_style = {};
  note_style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  note_style[DocumentApp.Attribute.FONT_FAMILY] = 'Google Sans';
  note_style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#cc0000';
  note_style[DocumentApp.Attribute.FONT_SIZE] = 10;
  note_style[DocumentApp.Attribute.BOLD] = false;
  note_style[DocumentApp.Attribute.SPACING_AFTER] = 12;
  note_style[DocumentApp.Attribute.SPACING_BEFORE] = 12;

  var bold_style = {};
  bold_style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  bold_style[DocumentApp.Attribute.FONT_FAMILY] = 'Google Sans';
  bold_style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#434343';
  bold_style[DocumentApp.Attribute.FONT_SIZE] = 10;
  bold_style[DocumentApp.Attribute.BOLD] = true;
  bold_style[DocumentApp.Attribute.SPACING_AFTER] = 12;
  bold_style[DocumentApp.Attribute.SPACING_BEFORE] = 12;
  
  
  // Open doc
  var doc_id = sheet.getRange(3, 3).getValue();
  emptyDocument(doc_id);
  var doc = DocumentApp.openById(doc_id);
  var body = doc.getBody();

  
  // Add title with heading 1
  var solution_title = sheet.getRange(2, 3).getValue();
  var solution_author = sheet.getRange(4, 3).getValue();
  var solution_level = sheet.getRange(5, 3).getValue();
  var solution_duration = sheet.getRange(6, 3).getValue();

  var author = body.insertParagraph(0, "Author: " + solution_author).setAttributes(author_style).setSpacingBefore(0).setSpacingAfter(0);
  var level = body.insertParagraph(1, "Level: " + solution_level)setAttributes(author_style).setSpacingBefore(0).setSpacingAfter(0);
  var time = body.insertParagraph(2, "Duration: "+ solution_duration).setAttributes(author_style).setSpacingBefore(0).setSpacingAfter(4);
  var title = body.insertParagraph(3, solution_title).setSpacingBefore(1).setSpacingAfter(0);
  var horizontalLine = body.insertHorizontalRule(4);
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  
  var listnum0;
  var listnum1;
    // For each row in the sheet, do the following
  for ( var row = 0; row < sheet_data.length; row++ ) {
    
    // If there is data in both Row1 (style) and Row2 (steps/headings)
    // Assign the appropriate heading settings
    if (sheet_data[row][0] && sheet_data[row][1]) {
      switch (sheet_data[row][0].toString()) {
        case 'H1':
          body.appendParagraph(sheet_data[row][1]).setHeading(DocumentApp.ParagraphHeading.HEADING1);
          break;
        case 'H2':
          body.appendParagraph(sheet_data[row][1]).setHeading(DocumentApp.ParagraphHeading.HEADING2);
          break;
        case 'H3':
          body.appendParagraph(sheet_data[row][1]).setHeading(DocumentApp.ParagraphHeading.HEADING3);
          break;
        case 'N':
          body.appendParagraph(sheet_data[row][1]).setAttributes(note_style);
          break;
        case '1':
          if (!listnum0) {  
            listnum0 = body.appendListItem(sheet_data[row][1]).setAttributes(steps_style).setGlyphType(DocumentApp.GlyphType.NUMBER);
            break;
          } else {
          listnum1 = body.appendListItem(sheet_data[row][1]).setAttributes(steps_style).setGlyphType(DocumentApp.GlyphType.NUMBER);
            listnum1.setListId(listnum0);
            break;
          }
        case '*':
          var listdot = body.appendListItem(sheet_data[row][1]).setAttributes(steps_style).setGlyphType(DocumentApp.GlyphType.BULLET);
          break;
        } 
    }
    
    // If only data in Row 2 (steps), then treat it as a normal paragraph and append
    if (!sheet_data[row][0] && sheet_data[row][1]) {
      body.appendParagraph(sheet_data[row][1]).setAttributes(steps_style);
    } 
    
    // If only data in Row 3 (code), then create a code block with the code_style attributes (indent, consolas etc)
    if (sheet_data[row][2]) {
      body.appendParagraph(sheet_data[row][2]).setAttributes(code_style);
    } 
    
    // If only data in Row 4 (output) then create an output block
    if (sheet_data[row][3]) {
      body.appendParagraph('\rOutput (do not copy)\r').setAttributes(output_style);
      body.appendParagraph(sheet_data[row][3]).setAttributes(output_style);
    }
    
  }
  
  // Get the body of the doc
  var body_data = doc.getBody();
  
  // Search for regex for text surrounded by double ticks (``, these are to the left of the numeral 1 key on the keyboard)
  var foundElement = body.findText('``.+?``');
  
  while (foundElement != null) {
    // Get the text object from the element
    var foundText = foundElement.getElement().asText();
    
    // Where in the Element is the found text?
    var start = foundElement.getStartOffset();
    var end = foundElement.getEndOffsetInclusive();

    // Set the style attribute to code_style
    foundText.setAttributes(start, end, code_style);

    // Find the next match
    foundElement = body.findText('``.+?``', foundElement);
    
  }
  
  // Remove double ticks after style change
  body_data.replaceText("``", ""); 

}

function emptyDocument(doc_id) {
  var document = DocumentApp.openById(doc_id);
  var body = document.getBody();
  body.appendParagraph('');// to be sure to delete the last paragraph in case it doesn't end with a cr/lf
  while (body.getNumChildren() > 1) body.removeChild( body.getChild( 0 ) );
}

// Create a custom menu option to be able to run the script from the Sheets app
// This function will create an adhoc list of everything in the Snippets sheet (sheetUrl) at the top of the Snippets doc (docUrl)
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Docs')
  .addItem('Generate Guide Template', 'createSolutionTemplate')
  .addToUi();
}
.
