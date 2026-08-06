function onOpen() {

  DocumentApp.getUi()
    .createMenu("Notation")

    .addItem("New Song", "showNewSongDialog")

    .addSeparator()

    .addItem("Octave Swaras", "showSwaraSidebar")

    .addSeparator()

    .addItem("Export Notation", "exportNotation")

    .addSeparator()

    .addItem("About", "showAbout")

    .addToUi();
}

function showNewSongDialog() {

  const html = HtmlService
    .createHtmlOutputFromFile("NewSongDialog")
    .setWidth(300)
    .setHeight(220);

  DocumentApp.getUi().showModalDialog(
    html,
    "New Song"
  );
}

function showSwaraSidebar() {

  const html = HtmlService
    .createHtmlOutputFromFile("SwaraSidebar")
    .setTitle("Octave Swaras");

  DocumentApp.getUi().showSidebar(html);
}


function insertSwara(swara) {

  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    return;
  }

  cursor.insertText(swara);
}

function showAbout() {
  DocumentApp.getUi().alert(
    "Hindustani Notation Editor\nVersion 1.0"
  );
}