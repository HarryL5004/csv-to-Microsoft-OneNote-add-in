/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import {generatePageHandler} from "./importCsv";
/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    
    // event listener for auto-generating pages from csv files
    const fileSelector: HTMLInputElement = document.getElementById("file-selector") as HTMLInputElement;
    fileSelector.addEventListener('change', generatePageHandler);
  }
});

export async function run() {
  try {
    await OneNote.run(async context => {
      // gets current page
      let page: OneNote.Page = context.application.getActivePage();

      // queue a command to set the page title.
      page.title = "Hello World";

      // queue a command to add an outline to the page.
      let html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
      page.addOutline(40, 90, html);

      // Run the queued commands, and return a promise to indicate task completion.
      return context.sync();
    });
  }
  catch(err) {
    console.log("Error: " + err);
  }
}
