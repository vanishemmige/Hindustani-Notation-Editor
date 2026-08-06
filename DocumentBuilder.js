
function createSongFromDialog(taalName) {

  const taal = TALAS[taalName];

  if (!taal) {
    DocumentApp.getUi().alert(
      "Taal not found: " + taalName
    );
    return;
  }

  createSong(taal);
}

function createSong(taal) {

  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    DocumentApp.getUi().alert(
      "Please place the cursor where you want to insert the song."
    );
    return;
  }

  let element = cursor.getElement();

  while (
    element &&
    element.getType() !== DocumentApp.ElementType.PARAGRAPH
  ) {
    element = element.getParent();
  }

  if (!element) {
    DocumentApp.getUi().alert(
      "Please place the cursor in a normal paragraph."
    );
    return;
  }

  const paragraph = element.asParagraph();
  const body = doc.getBody();

  let index = body.getChildIndex(paragraph) + 1;

  // -------------------------
  // SONG TITLE
  // -------------------------

  const songTitle = body.insertParagraph(index++, "Song Title");

  songTitle
    .setHeading(DocumentApp.ParagraphHeading.HEADING3)
    .setFontFamily(FONT);


  // -------------------------
  // SONG METADATA
  // -------------------------

  const metadata = [
    "Aaroh:",
    "Avaroh:",
    "Pakad:",
    "Taal: " + taal.name
  ];

  for (const line of metadata) {

    const p = body.insertParagraph(index++, line);

    p.setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .setFontFamily(FONT)
      .setFontSize(11)
      .setBold(false);
  }

  body.insertParagraph(index++, "");


  // -------------------------
  // ASTHAYI
  // -------------------------

  index = insertSectionAt(
    body,
    index,
    taal.title1,
    taal
  );

  body.insertParagraph(index++, "");


  // -------------------------
  // ANTRA
  // -------------------------

  index = insertSectionAt(
    body,
    index,
    taal.title2,
    taal
  );

  body.insertParagraph(index, "");
}


function insertSectionAt(body, index, title, taal) {

  // Section heading
  const heading = body.insertParagraph(index++, title);

  heading
    .setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setFontFamily(FONT)
    .setFontSize(14)
    .setBold(false);

  // Build one table row containing all vibhags
  const tableData = [[]];

  for (let i = 0; i < taal.markers.length; i++) {

    tableData[0].push(
      taal.markers[i] +
      "\n" +
      taal.matras[i] +
      "\n\n"
    );

  }

  // Insert table
  const table = body.insertTable(index++, tableData);

  // Set table font
  for (let r = 0; r < table.getNumRows(); r++) {

    const row = table.getRow(r);

    for (let c = 0; c < row.getNumCells(); c++) {

      row.getCell(c)
        .editAsText()
        .setFontFamily(FONT)
        .setFontSize(11)
        .setBold(false);

    }

  }

  return index;
}