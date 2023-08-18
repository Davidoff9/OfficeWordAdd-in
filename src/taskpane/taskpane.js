/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, console */

import { base64Image } from "../../base64Image";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle);
    document.getElementById("change-font").onclick = () => tryCatch(changeFont);
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange);
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextBeforeRange);
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText);
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
    document.getElementById("create-content-control").onclick = () => tryCatch(createContentControl);
    document.getElementById("replace-content-in-control").onclick = () => tryCatch(replaceContentInControl);
  }
});

async function replaceContentInControl() {
  await Word.run(async (context) => {
    //        content control.
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", Word.InsertLocation.replace);

    await context.sync();
  });
}

async function createContentControl() {
  await Word.run(async (context) => {
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";

    await context.sync();
  });
}

async function insertTable() {
  await Word.run(async (context) => {
    //        that will precede the table.
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

    const tableData = [
      ["Name", "ID", "Birth City"],
      ["David", "434", "Petrosani"],
      ["Intel", "719", "Core"],
    ];
    secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);

    await context.sync();
  });
}

async function insertHTML() {
  await Word.run(async (context) => {
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
    blankParagraph.insertHtml(
      '<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>',
      Word.InsertLocation.end
    );

    await context.sync();
  });
}

async function insertImage() {
  await Word.run(async (context) => {
    context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);

    await context.sync();
  });
}

async function replaceText() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", Word.InsertLocation.replace);

    await context.sync();
  });
}

async function insertTextBeforeRange() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    //        selected range.
    originalRange.insertText("Office 2019, ", Word.InsertLocation.before);

    //        range text can be read and inserted.
    originalRange.load("text");
    await context.sync();

    //        paragraph at the end of the document.
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, Word.InsertLocation.end);

    //        that it runs after the insertParagraph has been queued.
    await context.sync();
  });
}

async function insertTextIntoRange() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (M365)", Word.InsertLocation.end);

    //        current range text can be read.
    originalRange.load("text");
    await context.sync();

    //        range at the end of the document.
    doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);

    await context.sync();
  });
}

async function changeFont() {
  await Word.run(async (context) => {
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
    });

    await context.sync();
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";

    await context.sync();
  });
}

async function applyStyle() {
  await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;

    await context.sync();
  });
}

async function insertParagraph() {
  await Word.run(async (context) => {
    const docBody = context.document.body;
    docBody.insertParagraph(
      "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
      Word.InsertLocation.start
    );

    await context.sync();
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
