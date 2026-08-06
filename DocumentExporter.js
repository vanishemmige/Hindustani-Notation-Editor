function exportNotation() {

  const inputDoc = DocumentApp.getActiveDocument();
  const inputBody = inputDoc.getBody();

  const inputFile = DriveApp.getFileById(inputDoc.getId());
  const outputName = inputDoc.getName() + " - Export";

  // Get the folder containing the input document
  const parents = inputFile.getParents();

  if (!parents.hasNext()) {
    DocumentApp.getUi().alert(
      "Could not find the folder containing the input document."
    );
    return;
  }

  const folder = parents.next();
  Logger.log("Input folder: " + folder.getName());
  Logger.log("Folder ID: " + folder.getId());

  // Look for an existing output document
  const existingFiles = folder.getFilesByName(outputName);

  let outputDoc;

  if (existingFiles.hasNext()) {

    // Reuse existing output
    const outputFile = existingFiles.next();

    outputDoc = DocumentApp.openById(outputFile.getId());

    // Clear previous export
    outputDoc.getBody().clear();

  } else {

    // Create output document
    outputDoc = DocumentApp.create(outputName);

    const outputFile = DriveApp.getFileById(outputDoc.getId());

    // Move output to the same folder as input
    outputFile.moveTo(folder);
  }

  const outputBody = outputDoc.getBody();

  // Read the complete input document
  for (let i = 0; i < inputBody.getNumChildren(); i++) {

    const element = inputBody.getChild(i);

    // Ordinary text
  if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {

  const sourceParagraph = element.asParagraph();

  // If this is a song title (Heading 3),
  // start it on a new page — except for the first song.
  if (
    sourceParagraph.getHeading() ===
      DocumentApp.ParagraphHeading.HEADING3
  ) {

    // Don't add a page break if output is still empty
    if (outputBody.getText().trim() !== "") {
      outputBody.appendPageBreak();
    }
  }

  // Copy paragraph and preserve formatting
  outputBody.appendParagraph(
    sourceParagraph.copy()
  );
}

    // Notation table
    else if (element.getType() === DocumentApp.ElementType.TABLE) {

      exportNotationTable(
        element.asTable(),
        outputBody
      );
    }
  }

  outputDoc.saveAndClose();

  DocumentApp.getUi().alert(
    "Export complete.\n\n" +
    outputName
  );
}

function exportNotationTable(table, outputBody) {

  for (let r = 0; r < table.getNumRows(); r++) {

    const row = table.getRow(r);

    const cells = [];

    // Read every vibhag
    for (let c = 0; c < row.getNumCells(); c++) {

      const cell = row.getCell(c);

      cells.push({
        lines: cell.getText().split("\n"),
        formattedLines: getFormattedLines(cell)
      });
    }

    // getColumnWidths expects just the text lines
    const textCells = cells.map(cell => cell.lines);

    const widths = getColumnWidths(textCells);

    // Find maximum number of lines
    let maxLines = 0;

    for (const cell of cells) {
      maxLines = Math.max(maxLines, cell.lines.length);
    }

    // Export line by line
    for (let lineNo = 0; lineNo < maxLines; lineNo++) {

      let outputLine = "";

      // Remember exact underline positions in output
      const underlineRanges = [];

      for (let c = 0; c < cells.length; c++) {

        const formattedLine =
          cells[c].formattedLines[lineNo] || [];

        // Position where this vibhag begins
        const columnStart = outputLine.length;

        let cleanText = "";

        // Build this line token by token
        for (let t = 0; t < formattedLine.length; t++) {

          const token = formattedLine[t];

          if (t > 0) {
            cleanText += " ";
          }

          const tokenStart =
            columnStart + cleanText.length;

          cleanText += token.text;

          const tokenEnd =
            columnStart + cleanText.length - 1;

          if (token.underline) {

            underlineRanges.push({
              start: tokenStart,
              end: tokenEnd
            });

          }
        }

        // Add this vibhag to output
        if (c < cells.length - 1) {

          outputLine += pad(
            cleanText,
            widths[c]
          );

        } else {

          outputLine += cleanText;

        }
      }

      // Create plain-text paragraph
      const p = outputBody.appendParagraph(outputLine);

      p.setFontFamily(FONT);
      p.setFontSize(11);

      // Apply underline only to exact source swaras
      const outputText = p.editAsText();

      for (const range of underlineRanges) {

        outputText.setUnderline(
          range.start,
          range.end,
          true
        );

      }
    }

    // Blank line after notation row
    outputBody.appendParagraph("");
  }
}

function getFormattedLines(cell) {

  const text = cell.editAsText();
  const fullText = text.getText();

  const lines = [];
  let currentLine = [];

  // Find tokens and remember their exact formatting
  const regex = /\S+/g;
  let match;

  while ((match = regex.exec(fullText)) !== null) {

    const token = match[0];
    const start = match.index;
    const end = start + token.length - 1;

    // If we crossed a newline, start a new line
    const textBeforeToken = fullText.substring(0, start);
    const lineNumber =
      (textBeforeToken.match(/\n/g) || []).length;

    while (lines.length <= lineNumber) {
      lines.push([]);
    }

    let underlined = false;

    // Check this exact occurrence
    for (let i = start; i <= end; i++) {

      if (text.isUnderline(i)) {
        underlined = true;
        break;
      }
    }

    lines[lineNumber].push({
      text: token,
      underline: underlined
    });
  }

  return lines;
}

function pad(text, width) {

  while (text.length < width) {
    text += " ";
  }

  return text;
}

function getColumnWidths(cells) {

  const widths = [];

  for (let c = 0; c < cells.length; c++) {

    let max = 0;

    for (const line of cells[c]) {

      const text = line.trim().replace(/\s+/g, " ");

      if (text.length > max) {
        max = text.length;
      }

    }

    widths.push(max + 6);   // 6 spaces between vibhags

  }

  return widths;

}

function getUnderlinedTokens(cell) {

  const text = cell.editAsText();
  const fullText = text.getText();

  const underlined = new Set();

  // Find every non-space token
  const regex = /\S+/g;
  let match;

  while ((match = regex.exec(fullText)) !== null) {

    const token = match[0];
    const start = match.index;
    const end = start + token.length - 1;

    // Consider the token underlined if any character
    // in that token is underlined
    let isUnderlined = false;

    for (let i = start; i <= end; i++) {

      if (text.isUnderline(i)) {
        isUnderlined = true;
        break;
      }
    }

    if (isUnderlined) {
      underlined.add(token);
    }
  }

  return underlined;
}
