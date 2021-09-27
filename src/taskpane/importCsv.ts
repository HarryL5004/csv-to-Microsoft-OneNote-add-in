// main functionality for importing csv

// separate a record into several components (based on csv schema)
function splitEntry(entry: string) {
    // custom function; should be tailored to your own needs
}

// generate HTML string for each record
// returns: [[string, HTML string], ...]
function generateHTMLFromFile(fileContent: string) {
    let entries: Array<string> = fileContent.split(new RegExp('\r\n|[\n\r]'));

    return entries.map(entry => {
        let components: Array<string> = splitEntry(entry);
        let date: string = components[0];
        let content: string = components[1];

        return [date, `<p>${content}</p>`]
    });
}

// create a OneNote page for each record
async function createPages(entries: Array<Array<string>>) {
    try {
        await OneNote.run(async context => {
        // get active Section
        let section: OneNote.Section = context.application.getActiveSection();
        // for each record create a new page with the content
        entries.forEach(entry => {
            // insert new page into section
            let newPage = section.addPage(entry[0]);
            // set active page to newly created page
            context.application.navigateToPage(newPage);
            // create an outline to put text in
            newPage.addOutline(40, 90, entry[1]);
        });
        
        return context.sync();
        });
    } catch (e) {
        console.log("Error: " + e);
    }
}

// event handler for generating pages from csv files
export function generatePageHandler(e) {
    let file = (e.target as HTMLInputElement).files[0];
    let reader = new FileReader();
  
    reader.addEventListener('load', (e) => {
      console.log("Loaded File");
  
      let fileContent = e.target.result as string;
      let entries: Array<Array<string>> = generateHTMLFromFile(fileContent);
      createPages(entries);
    });
    reader.readAsText(file);
 }