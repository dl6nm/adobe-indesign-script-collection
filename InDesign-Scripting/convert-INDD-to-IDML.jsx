// Open an InDesign (indd) file using JavaScript, ignore missing links, and save as IDML.

// Set parameters
var showDocumentInWindow = true;  // Set to true to open the InDesign file in a window.
var logFile = new File(Folder.myDocuments + "/indd-to-idml.log");
var updateLinks = false;  // Set to true to update missing links.

// Set script preferences
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
app.scriptPreferences.enableRedraw = true;

try {
    // Init logger
    var logger = getLogger(logFile);

    // Open an InDesign file
    var file = File.openDialog("Select the InDesign file.", "*.indd");

    // Convert the InDesign file to IDML
    convertInddToIdml(file);
} catch (e) {
    log(e);
} finally {
    // Close the log file.
    closeLogger(logFile);
}

function convertInddToIdml(file) {
    // Opens an InDesign INDD file and exports it as IDML.
    try {
        // Open file in InDesign
        var doc = app.open(file, showDocumentInWindow);

        // Update or ignore missing links.
        if (updateLinks) {
            doc.links.everyItem().update();
        }
        
        // Define the IDML file.
        var idmlFile = new File(file.path + "/" + doc.name.replace(/\.indd$/, ".idml"));

        // Save the IDML file.
        doc.exportFile(ExportFormat.INDESIGN_MARKUP, idmlFile);
    } catch (e) {
        log(e);
    } finally {
        // Close the InDesign file
        doc.close(SaveOptions.NO);
    }
}

function getLogger(logFile) {
    logFile.encoding = "UTF-8";
    return logFile.open("a");
}

function closeLogger(logFile) {
    logFile.close();
}

function log(message) {
    logger.writeln(message);
}
