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
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Bologna Law Review');
  DocumentApp.getUi().showSidebar(ui);
}


/*************************************************************
* CONVERSION ENGINE FUNCTIONS
*************************************************************/

/**
 * LaTeX conversion engine, takes the whole document and formats the style
 * to a suitable standard for TeX handling, based on the great doc2md 
 * converter by Renato Mangini https://github.com/mangini/gdocs2md.git
 */

function ConvertToLatex() {
    
  var numChildren = DocumentApp.getActiveDocument().getActiveSection().getNumChildren();
  var text = "";
  var inSrc = false;
  var inClass = false;
  var globalImageCounter = 0;
  var globalListCounters = {};
  // edbacher: added a variable for indent in src <pre> block. Let style sheet do margin.
  var srcIndent = "";
  
  var attachments = [];
  
  // Walk through all the child elements of the doc.
  for (var i = 0; i < numChildren; i++) {
      
    var child = DocumentApp.getActiveDocument().getActiveSection().getChild(i);
   
    var result = processParagraph(i, child, inSrc, globalImageCounter, globalListCounters);
    
    globalImageCounter += (result && result.images) ? result.images.length : 0;
    
    if (result!==null) {
      if (result.sourcePretty==="start" && !inSrc) {
        inSrc=true;
        text+="<pre class=\"prettyprint\">\n";
      } else if (result.sourcePretty==="end" && inSrc) {
        inSrc=false;
        text+="</pre>\n\n";
      } else if (result.source==="start" && !inSrc) {
        inSrc=true;
        text+="<pre>\n";
      } else if (result.source==="end" && inSrc) {
        inSrc=false;
        text+="</pre>\n\n";
      } else if (result.inClass==="start" && !inClass) {
        inClass=true;
        text+="<div class=\""+result.className+"\">\n";
      } else if (result.inClass==="end" && inClass) {
        inClass=false;
        text+="</div>\n\n";
      } else if (inClass) {
        text+=result.text+"\n\n";
      } else if (inSrc) {
        text+=(srcIndent+escapeHTML(result.text)+"\n");
      } else if (result.text && result.text.length>0) {
        text+=result.text+"\n\n";
      }
      
      if (result.images && result.images.length>0) {
        for (var j=0; j<result.images.length; j++) {
          attachments.push( {
            "fileName": result.images[j].name,
            "mimeType": result.images[j].type,
            "content": result.images[j].bytes } );
        }
      }
    } else if (inSrc) { // support empty lines inside source code
      text+='\n';
    }
      
  }

  // Return the text to be displayed on the text box
  return {
    text: text
  };
}

function escapeHTML(text) {
  return text.replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Process each child element (not just paragraphs).
function processParagraph(index, element, inSrc, imageCounter, listCounters) {
  // First, check for things that require no processing.
  if (!element.getNumChildren) {
    // just in case we would cut something whose absence wouldn't be obvious...
    return {"text": "% There was something here that couldn't be converted."};
  }
  if (element.getNumChildren()==0) {
    return null;
  }  
  // Punt on TOC.
  if (element.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
    return {"text": "[[TOC]]"};
  }
  
  // Set up for real results.
  var result = {};
  var pOut = "";
  var textElements = [];
  var imagePrefix = "image_";
  
  // Handle Table elements. Pretty simple-minded now, but works for simple tables.
  // Note that Markdown does not process within block-level HTML, so it probably 
  // doesn't make sense to add markup within tables.
  if (element.getType() === DocumentApp.ElementType.TABLE) {
    textElements.push("<table>\n");
    var nCols = element.getChild(0).getNumCells();
    for (var i = 0; i < element.getNumChildren(); i++) {
      textElements.push("  <tr>\n");
      // process this row
      for (var j = 0; j < nCols; j++) {
        textElements.push("    <td>" + element.getChild(i).getChild(j).getText() + "</td>\n");
      }
      textElements.push("  </tr>\n");
    }
    textElements.push("</table>\n");
  }
  
  // Process various types (ElementType).
  for (var i = 0; i < element.getNumChildren(); i++) {
    var t=element.getChild(i).getType();
  
    if (t === DocumentApp.ElementType.TABLE_ROW) {
      // TABLE ROW (ALREADY HANDLED)
      
    } else if (t === DocumentApp.ElementType.TEXT) {
      // SIMPLE TEXT
      var txt=element.getChild(i);
      pOut += txt.getText();
      textElements.push(txt);
      
    } else if (t === DocumentApp.ElementType.INLINE_IMAGE) {
      // IMAGES
      result.images = result.images || [];
      var contentType = element.getChild(i).getBlob().getContentType();
      var extension = "";
      if (/\/png$/.test(contentType)) {
        extension = ".png";
      } else if (/\/gif$/.test(contentType)) {
        extension = ".gif";
      } else if (/\/jpe?g$/.test(contentType)) {
        extension = ".jpg";
      } else {
        throw "Unsupported image type: "+contentType;
      }
      var name = imagePrefix + imageCounter + extension;
      imageCounter++;
      textElements.push('![image alt text]('+name+')');
      result.images.push( {
        "bytes": element.getChild(i).getBlob().getBytes(), 
        "type": contentType, 
        "name": name});
      
    } else if (t === DocumentApp.ElementType.PAGE_BREAK) {
      // IGNORE PAGE BREAKS
    
    } else if (t === DocumentApp.ElementType.HORIZONTAL_RULE) {  
      // HORIZONTAL RULES
      textElements.push('\\horrule');
    
    } else if (t === DocumentApp.ElementType.FOOTNOTE) {  
      // FOOTNOTES
      var note = processTextElement(inSrc, element.getChild(i).getFootnoteContents().asText());
      textElements.push('\\footnote{'+note+'}');
    
    } else {  
      // UNSUPPORTED
      throw "Paragraph "+index+" of type "+element.getType()+" has an unsupported child: "
      +t+" "+(element.getChild(i)["getText"] ? element.getChild(i).getText():'')+" index="+index;
    }
  }

  if (textElements.length==0) {
    // Isn't result empty now?
    return result;
  }
  
  // evb: Add source pretty too. (And abbreviations: src and srcp.)
  // process source code block:
  if (/^\s*---\s+srcp\s*$/.test(pOut) || /^\s*---\s+source pretty\s*$/.test(pOut)) {
    result.sourcePretty = "start";
  } else if (/^\s*---\s+src\s*$/.test(pOut) || /^\s*---\s+source code\s*$/.test(pOut)) {
    result.source = "start";
  } else if (/^\s*---\s+class\s+([^ ]+)\s*$/.test(pOut)) {
    result.inClass = "start";
    result.className = RegExp.$1;
  } else if (/^\s*---\s*$/.test(pOut)) {
    result.source = "end";
    result.sourcePretty = "end";
    result.inClass = "end";
  } else if (/^\s*---\s+jsperf\s*([^ ]+)\s*$/.test(pOut)) {
    result.text = '<iframe style="width: 100%; height: 340px; overflow: hidden; border: 0;" '+
                  'src="http://www.html5rocks.com/static/jsperfview/embed.html?id='+RegExp.$1+
                  '"></iframe>';
  } else {

    prefix = findPrefix(inSrc, element, listCounters);
  
    var pOut = "";
    for (var i=0; i<textElements.length; i++) {
      pOut += processTextElement(inSrc, textElements[i]);
    }

    // replace Unicode quotation marks
    pOut = pOut.replace('\u201d', '"').replace('\u201c', '"');
    
    result.text = prefix[0]+pOut+prefix[1];
  }
  
  return result;
}

// Add correct prefix to list items.
function findPrefix(inSrc, element, listCounters) {
  var prefix="";
  var suffix="";
  if (!inSrc) {
    if (element.getType()===DocumentApp.ElementType.PARAGRAPH) {
      // SECTIONS
      var paragraphObj = element;
      switch (paragraphObj.getHeading()) {
        // Each heading corresponds to a section type
        case DocumentApp.ParagraphHeading.HEADING1: 
          prefix+="\\section{";
          suffix+="}";
          break;
        case DocumentApp.ParagraphHeading.HEADING2: 
          prefix+="\\subsection{";
          suffix+="}";
          break;
        case DocumentApp.ParagraphHeading.HEADING3: 
          prefix+="\\subsubsection{";
          suffix+="}";
          break;
        default:
      }
    } else if (element.getType()===DocumentApp.ElementType.LIST_ITEM) {
      // LIST ITEM
      var listItem = element;
      var nesting = listItem.getNestingLevel()
      for (var i=0; i<nesting; i++) {
        prefix += "    ";
      }
      var gt = listItem.getGlyphType();
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET || gt === DocumentApp.GlyphType.HOLLOW_BULLET || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        // Test to see if the first item is reached (or if previous siblings exist)
        try{
          if (listItem.getPreviousSibling().getType() !== listItem.getType()){
            prefix += "\\begin{itemize}\n";
          } else if (listItem.getNestingLevel() > listItem.getPreviousSibling().getNestingLevel()){
            // Add subsequent levels of indentetions
            prefix += "\\begin{itemize}\n";
          }
        }
        catch (err){
          // If the sibling does not exist, the item is the first
          prefix += "\\begin{itemize}\n";
        }
        prefix += "\\item ";
        // Test to see if the last item is reached (or if next siblings exist)
        try{
          if (listItem.getNextSibling().getType() !== listItem.getType()){
            suffix += "\n\\end{itemize}\n";
          } else if (listItem.getNestingLevel() > listItem.getNextSibling().getNestingLevel()){
            // Add subsequent levels of indentetions
            for (var i = listItem.getNestingLevel(); i > listItem.getNextSibling().getNestingLevel(); i--)
              suffix += "\n\\end{itemize}\n";
          }
        } catch (err){
          // If the sibling does not exist, the item is the last
          for (var i = listItem.getNestingLevel(); i >= 0; i--)
              suffix += "\n\\end{itemize}\n";
        }
      } else {
        // Ordered list (<ol>):
        var key = listItem.getListId() + '.' + listItem.getNestingLevel();
        var counter = listCounters[key] || 0;
        counter++;
        listCounters[key] = counter;
        // Test to see if the first item is reached (or if previous siblings exist)
        try {
          if (listItem.getPreviousSibling().getType() !== listItem.getType()){
            prefix += "\\begin{enumerate}\n";
          } else if (listItem.getNestingLevel() > listItem.getPreviousSibling().getNestingLevel()){
            // Add subsequent levels of indentetions
            prefix += "\\begin{enumerate}\n";
          }
        }
        catch(err){
          // If the sibling does not exist, the item is the first
          prefix += "\\begin{enumerate}\n";
        }
        // Add the standard item prefix
        prefix += "\\item ";
        // Test to see if the last item is reached (or if next siblings exist)
        try{
          if (listItem.getNextSibling().getType() !== listItem.getType()){
            suffix += "\n\\end{enumerate}\n";
          } else if (listItem.getNestingLevel() > listItem.getNextSibling().getNestingLevel()){
            // Add subsequent levels of indentetions
            for (var i = listItem.getNestingLevel(); i > listItem.getNextSibling().getNestingLevel(); i--)
              suffix += "\n\\end{enumerate}\n";
          }
        }
        catch(err){
          // If the sibling does not exist, the item is the last
          for (var i = listItem.getNestingLevel(); i >= 0; i--)
              suffix += "\n\\end{enumerate}\n";
        }
      }
    }
  }
  return [prefix, suffix];
}

function processTextElement(inSrc, txt) {
  if (typeof(txt) === 'string') {
    return txt;
  }
  
  var pOut = txt.getText();
  if (! txt.getTextAttributeIndices) {
    return pOut;
  }
  
  var attrs=txt.getTextAttributeIndices();
  var lastOff=pOut.length;

  for (var i=attrs.length-1; i>=0; i--) {
    var off=attrs[i];

    var url=txt.getLinkUrl(off);
    var font=txt.getFontFamily(off);
    
    if (url) {  // start of link
      if (i>=1 && attrs[i-1]==off-1 && txt.getLinkUrl(attrs[i-1])===url) {
        // detect links that are in multiple pieces because of errors on formatting:
        i-=1;
        off=attrs[i];
        url=txt.getLinkUrl(off);
      }
      pOut=pOut.substring(0, off)+'\\href{'+url+'}{'+pOut.substring(off, lastOff)+'}'+pOut.substring(lastOff);
      
    } else if (font) {
      if (!inSrc && font===font.COURIER_NEW) {
        while (i>=1 && txt.getFontFamily(attrs[i-1]) && txt.getFontFamily(attrs[i-1])===font.COURIER_NEW) {
          // detect fonts that are in multiple pieces because of errors on formatting:
          i-=1;
          off=attrs[i];
        }
        pOut=pOut.substring(0, off)+'\\verbatim{'+pOut.substring(off, lastOff)+'}'+pOut.substring(lastOff);
      }
    }
    
    if (txt.isBold(off) || txt.isItalic(off) || txt.isUnderline(off) || txt.isStrikethrough(off)){
      var d1 = ""; d2 = "";
      if (txt.isBold(off)) {
        // BOLD
        d1 += "\\textbf{"; d2 += "}";
      }
      if (txt.isItalic(off)) {
        // ITALIC
        d1 += "\\textit{"; d2 += "}"; 
      }
      if (txt.isUnderline(off)) {
        // UNDERLINED
        d1 += "\\underline{"; d2 += "}";
      }
      if (txt.isStrikethrough(off)) {
        // STRIKETROUGH
        d1 += "\\striketrhough{"; d2 += "}";
      }
      pOut=pOut.substring(0, off)+d1+pOut.substring(off, lastOff)+d2+pOut.substring(lastOff);
    }
    
    if (off > 0){
      // SMALL CAPS
      if (txt.getType() === DocumentApp.ElementType.FOOTNOTE_SECTION) {
        if (txt.getFontSize(off-1) > txt.getFontSize(off)){
          // FOR FOOTNOTES
          var d1 = "\\textsc{"; d2 = "}";
          // Convert only if the string is actually fake small caps
          if (pOut.substring(off, lastOff).toUpperCase() == pOut.substring(off, lastOff)){
            if (pOut[off-1] == ' '){
              pOut=pOut.substring(0, off)+d1+pOut.substring(off, lastOff).toLowerCase()+d2+pOut.substring(lastOff);
            } else {
              pOut=pOut.substring(0, off-1)+d1+pOut.substring(off-1, off)+pOut.substring(off, lastOff).toLowerCase()+d2+pOut.substring(lastOff);
            }
          }
        }
      } else {
        if (txt.getFontSize(off-1) < txt.getFontSize(off)){
          // FOR NORMAL TEXT
          var d1 = "\\textsc{"; d2 = "}";
          // Convert only if the string is actually fake small caps
          if (pOut.substring(off, lastOff).toUpperCase() == pOut.substring(off, lastOff)){
            if (pOut[off-1] == ' '){
              pOut=pOut.substring(0, off)+d1+pOut.substring(off, lastOff).toLowerCase()+d2+pOut.substring(lastOff);
            } else {
              pOut=pOut.substring(0, off-1)+d1+pOut.substring(off-1, off)+pOut.substring(off, lastOff).toLowerCase()+d2+pOut.substring(lastOff);
            }
          }
        }
      }
    }
    lastOff=off;
  }
  
  pOut = processLatex(pOut);
  pOut = processLanguages(pOut);
  return pOut;
}


/**
 * Processes text to escape all special LaTeX characters
 * NOTE: needs \usepackage{eurosym}
 */
function processLatex(pOut){
  // Escape all special characters 
  for (c in latexEscapeCharacters)
   pOut = pOut.split(c).join(latexEscapeCharacters[c]);
  return pOut;
}

var latexEscapeCharacters = { 
  '%' : '\\%',
  '&' : '\\&',
  '$' : '\\$',
  '_' : '\\_',
  '#' : '\\#',
  '€' : '\\euro',
  '’' : '\''
};

/**
 * Processes text to add commands for different characters 
 * encoding
 * NOTE: needs \usepackage[greek,english]{babel}
 * NOTE: greek chars are between 0x0370 (880) and 0x03FF (1023)
 */
function processLanguages(pOut) {
  
  txt = pOut; 
  found = false;
  var begin = 0;
  var offset = 0;
  
  for (var count=0; count < txt.length; count++){
    
    // Detect the first greek caps char
    if (txt.charCodeAt(count) >= 880 && txt.charCodeAt(count) <= 1023){
      // Set "found" flag to start searching loop
      found = true;
      begin = count
    }
    
    while (found){
      // If the char is greek or a space keep counting up
      if ((txt.charCodeAt(count) >= 880 && txt.charCodeAt(count) <= 1023) || txt.charCodeAt(count) === ' '.charCodeAt(0)){
        count++;
      }
      else {
        // Reset "found" flag to break searching loop
        found = false;
        message = "\\textgreek{" + pOut.substring(begin+offset-1, count+offset) + "}";
        pOut=pOut.substring(0, begin+offset-1) + message + pOut.substring(count+offset);
        offset += message.length - pOut.substring(begin+offset-1, count+offset).length;
        count ++;
      }
    }
  }
  
  return pOut;
}


/*************************************************************
* CAPITALIZATION FUNCTIONS
*************************************************************/

/**
 * Takes selected text as input and creates formatted
 * output faking small caps by resizing lowercase letters 
 * to be uppercase with a smaller font or by reversing
 * the process to create normal caps based on "mode"
 *
 * @(mode) "smallcaps" or "normalcaps" 
 */

function textCapitalize(mode){
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var elements = selection.getSelectedElements();
    // Check all the elements sequentially
    for (var i = 0; i < elements.length; ++i) {
      // Discriminate between partial or selection
      if (elements[i].isPartial()) {
        // Get text with relative start and end
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        // Select conversion based on mode
        switch (mode){
          case "smallcaps":
            // Convert to small caps from start (startIndex) to end (endIndex)
            textToSmallCaps (element, startIndex, endIndex);
            break;
          case "normalcaps":
            // Convert to small caps from start (startIndex) to end (endIndex)
            textToNormalCaps (element, startIndex, endIndex);
            break;
        }
      } else {
        var element = elements[i].getElement();
        // Only translate elements that can be edited as text; skip images and
        // other non-text elements.
        if (element.editAsText) {
          var elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText) {
            // Sample formatted text and string content
            var txt = element.asText();
            // Select conversion based on mode
            switch (mode){
              case "smallcaps":
                // Convert to small caps from start (0) to end (length)
                textToSmallCaps (txt, 0, txt.getText().length);
                break;
              case "normalcaps":
                // Convert to small caps from start (0) to end (length)
                textToNormalCaps (txt, 0, txt.getText().length);
                break;
            }
          }
        }
      }
    }
  }
}

/**
 * Takes selected text as input and creates formatted
 * small caps as output faking small caps by resizing
 * lowercase letters to be uppercase with a smaller font 
 */

function textToSmallCaps (txt, startIndex, endIndex){
  var offset = 0;
  var pOut = txt.getText().substring(startIndex, endIndex + 1); 
  for (var i = 0; i < pOut.length; i++){ 
    
    if (pOut[i] == pOut.charAt(i).toUpperCase()){
      if (offset != i){
        var str = pOut.substring(offset, i);
        var size = txt.getFontSize(startIndex+offset-(offset!=startIndex));
        txt.insertText(startIndex+offset, str.toUpperCase());
        txt.setFontSize(startIndex+offset, startIndex+i, size*0.77 + 9*(!size));
        txt.deleteText(startIndex+offset+str.length, startIndex+offset+2*str.length-1)
      }
      offset = i+1;
    }
  }
  // Create string to substitute
  var str = pOut.substring(offset, i);
  // Check if the text needs to be converted
  if (startIndex+offset+2*str.length-1 > endIndex){
    var size = txt.getFontSize(startIndex+offset-(offset!=startIndex));
    txt.insertText(startIndex+offset, str.toUpperCase());
    txt.setFontSize(startIndex+offset, startIndex+i, size*0.77 + 9*(!size));
    txt.deleteText(startIndex+offset+str.length, startIndex+offset+2*str.length-1);
  } else {
    throw "Text is already small caps"
  }
}

/** TODO: Experimental new feature
 * Takes selected text as input and creates formatted
 * normale caps as output by resizing smaller font 
 * letters
 */

//function textToNormalCaps (txt, startIndex, endIndex){
//  var offset = 0;
//  var pOut = txt.getText().substring(startIndex, endIndex + 1); 
//  for (var i = 0; i < pOut.length; i++){ 
//    // Detect variations in form
//    if (txt.getFontSize(startIndex+i+1) < txt.getFontSize(startIndex+i)){
//      if (offset != i){
//        var str = pOut.substring(offset+1, i+1);
//        var size = txt.getFontSize(startIndex+offset-(offset!=startIndex));
//        txt.insertText(startIndex+offset+1, str.toLowerCase());
//        txt.deleteText(startIndex+offset+1+str.length, startIndex+offset+2*str.length)
//      }
//      offset = i+2;
//    }
//  }
//}

