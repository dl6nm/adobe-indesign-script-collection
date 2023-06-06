// Open an InDesign (indd) file using JavaScript, ignore missing links, and save as IDML.

// Set default parameters
var showDocumentInWindow = true; // Set to true to open the InDesign file in a window.
var logFile = new File(Folder.myDocuments + "/indd-to-idml.log");
var updateLinks = true; // Set to true to update missing links.
var exportPdf = true; // Set to true to export a PDF file.

// Set script preferences
app.scriptPreferences.userInteractionLevel =
    UserInteractionLevels.NEVER_INTERACT;
app.scriptPreferences.enableRedraw = true;

/*
    #########################################
    # Main Script
    #########################################
*/
try {
    // Init logger
    var logger = getLogger(logFile);

    // Ask the user to confirm a recursive conversion of InDesign files to IDML, otherwise convert just a single file.
    var recursive = confirm(
        "Convert all InDesign files from a folder and its subfolders to IDML? Otherwise, convert just a single file.",
        false,
        "Convert InDesign files to IDML"
    );

    // Ask the user to confirm also the export of a PDF file.
    var exportPdf = confirm(
        "Export also a PDF file?",
        false,
        "Export PDF file"
    );

    if (!recursive) {
        // Convert a single file.
        // Open an InDesign file
        var file = File.openDialog("Select the InDesign file.", "*.indd");
        // Convert the InDesign file to IDML and optionally export a PDF file.
        if (file) {
            convertInddToIdml(file);
        }
    } else {
        // Convert a folder and its subfolders.
        // Scan a folder and its subfolders for InDesign files and convert them to IDML.
        var folder = Folder.selectDialog("Select the folder to scan.");
        recursiveConvertInddToIdml(folder);
    }
} catch (e) {
    log(e);
} finally {
    // Close the log file.
    closeLogger(logFile);
}

/*
    #########################################
    # Functions
    #########################################
*/

function recursiveConvertInddToIdml(folder) {
    /**
     * Recursively scans a folder an its subfolders for InDesign files and converts them to IDML.
     * @param {Folder} folder - The folder to scan.
     */
    // Convert InDesign files to IDML.
    var inddFiles = folder.getFiles("*.indd");
    for (var i = 0; i < inddFiles.length; i++) {
        convertInddToIdml(inddFiles[i]);
    }

    // Scan subfolders.
    var subFolders = folder.getFiles("*");
    for (var i = 0; i < subFolders.length; i++) {
        if (subFolders[i] instanceof Folder) {
            recursiveConvertInddToIdml(subFolders[i]);
        }
    }
}

function convertInddToIdml(file) {
    /**
     * Opens an InDesign INDD file and exports it as IDML.
     * @param {File} file - The InDesign file to open.
     */
    try {
        // Open file in InDesign
        var doc = app.open(file, showDocumentInWindow);

        // Update or ignore missing links.
        if (updateLinks) {
            try {
                var links = doc.links;
                for (var i = 0; i < links.length; i++) {
                    var link = links[i];
                    if (link.status == LinkStatus.LINK_OUT_OF_DATE) {
                        link.update();
                    } else if (link.status == LinkStatus.LINK_MISSING) {
                        log("Link missing: " + link.filePath);
                    } 
                }
            } catch (e) {
                log(e);
            }
        }

        // Export a PDF file.
        if (exportPdf) {
            exportPdfFile(doc);
        }

        // Define the IDML file.
        var idmlFile = new File(
            file.path + "/" + doc.name.replace(/\.indd$/, ".idml")
        );

        // Save the IDML file.
        doc.exportFile(ExportFormat.INDESIGN_MARKUP, idmlFile);
    } catch (e) {
        log(e);
    } finally {
        // Close the InDesign file
        doc.close(SaveOptions.NO);
    }
}

function exportPdfFile(doc) {
    /**
     * Exports an open InDesign file as PDF.
     * @param {Document} doc - The InDesign document to export.
     */
    try {
        // Define the PDF file.
        var pdfFile = new File(
            file.path + "/" + file.name.replace(/\.indd$/, "_preview.pdf")
        );

        // Export the PDF file.
        doc.exportFile(ExportFormat.PDF_TYPE, pdfFile);
    } catch (e) {
        log(e);
    }
}

function getLogger(logFile) {
    logFile.encoding = "UTF-8";
    return logFile.open("a");
}

function closeLogger(logFile) {
    /**
     * Closes the log file.
     * @param {File} logFile - The log file to close.
     */
    logFile.close();
}

function log(message) {
    /**
     * Writes a message to the log file.
     * @param {string} message - The message to write to the log file.
     */
    logger.writeln(message);
}
