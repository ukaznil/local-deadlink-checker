@if(0) == (0) echo off
REM ::::::::::::::::::::::::::::::::::::::::::
REM ::                                      ::
REM ::    Windows batch part to operate     ::
REM ::    the following JScript module.     ::
REM ::                                      ::
REM ::::::::::::::::::::::::::::::::::::::::::
cscript.exe //nologo //E:JScript "%~f0" %*
@pause
goto :EOF
@end


//////////////////////////////////////////
//                                      //
//    JScript functions for             //
//    the following main procedure.     //
//                                      //
//////////////////////////////////////////

/**
 * echo String on a dialog.
 */
function echo(str) {
    WScript.echo(str);
}

/**
 * output String to stdout.
 */
function stdout(str) {
    var str = typeof str !== 'undefined' ? str : "";
    WScript.StdOut.writeLine("" + str);
}

/**
 * backport of 'startsWith' implementation for old JavaScript.
 */
if (!String.prototype.startsWith) {
    String.prototype.startsWith = function(prefix) {
        return this.lastIndexOf(prefix, 0) === 0;
    }
}

/**
 * backport of 'endsWith' implementation for old JavaScript.
 */
if (!String.prototype.endsWith) {
    String.prototype.endsWith = function(suffix) {
        return this.indexOf(suffix, this.length - suffix.length) !== -1;
    }
}

/**
 * backport of 'includes' implementation for old JavaScript.
 */
if (!String.prototype.includes) {
    String.prototype.includes = function(searchString) {
        return this.indexOf(searchString) !== -1
    }
}

/**
 * get all the files under the directory recursively.
 */
function enumFiles(dirPath) {
    var validExts = [".html", ".htm"];

    if (fso.fileExists(dirPath)) {
        return [dierPath];
    }

    if (!fso.folderExists(dirPath)) {
        return [];
    }

    var result = [];
    var dir = fso.getFolder(dirPath);

    // proceed to all files.
    var e1 = new Enumerator(dir.files);
    for (e1.moveFirst(); !e1.atEnd(); e1.moveNext()) {
        var file = e1.item();

        // check ext.
        for (var i in validExts) {
            if (file.path.endsWith(validExts[i])) {
                result.push(file.path);
                break;
            }
        }
    }

    // proceed to all dirs.
    var e2 = new Enumerator(dir.subFolders);
    for (e2.moveFirst(); !e2.atEnd(); e2.moveNext()) {
        var files = enumFiles(e2.item());
        result = result.concat(files);
    }

    return result;
}

/**
 * read all text from a file.
 */
function readFile(filePath) {
    // iomode
    var ForReading = 1;
    var ForWriting = 2;
    var ForAppending = 8;

    // format
    var TristateTrue = -1; // Unicode
    var TristateFalse = 0; // ASCII
    var TristateUseDefault = -2; // System Default

    // read file
    var file = fso.openTextFile(filePath, ForReading, create = false, TristateUseDefault);

    // empty file or not
    if (file.atEndofStream) {
        var text = "";
    } else {
        var text = file.readAll();
    }

    file.close();
    return text;
}

/**
 * get all the links written in the file.
 */
function checkAllLinks(filePath) {
    var wholeText = readFile(filePath);
    var document = new ActiveXObject("htmlfile");
    document.write(wholeText);

    var wholeDocument = document.getElementsByTagName("html")[0];

    function checkLinksByTag(tagName, attrName) {
        var elements = document.getElementsByTagName(tagName);
        for (var i = 0, len = elements.length; i < len; i++) {
            var element = elements[i];

            var regex = new RegExp(attrName + "=\"([^\"]*)");
            var matched = element.outerHTML.match(regex);

            if (matched !== null){
                var link = matched[1].replace(/&amp;/g, "&");

                var status = getLinkStatus(filePath, link);
                var msg = (status === "NG" ? "  !!" : "    ")
                    + "[" + status + "] " + element.outerHTML.replace(/&amp;/g, "&");

                // output to stdout.
                stdout(msg);

                // output to corresponding files.
                switch(status) {
                case "OK":
                    appendToFile(outputFileForOK.path, msg);
                    break;
                case "NG":
                    appendToFile(outputFileForNG.path, msg);
                    break;
                case "--":
                    appendToFile(outputFileForWWW.path, msg);
                    break;
                default:
                    echo("Unknown error.");
                    destroyObjects();
                    WScript.quit();
                    break;
                }
            }
        }
    }

    checkLinksByTag("a", "href");
    checkLinksByTag("img", "src");

    document = null;
}

/**
 * check whether the link is dead or alive.
 */
function getLinkStatus(filePath, link) {
    if (link.startsWith("http://") ||
        link.startsWith("https://") ||
        link.startsWith("//") ||
        link.startsWith("mailto:")) {
        var status = "--";
    } else {
        var path = fso.buildPath(fso.getParentFolderName(filePath), link);
        var status = fso.fileExists(path) ? "OK" : "NG";
    }

    return status;
}

/**
 * format int value to 2-digit String.
 */
function format2d(value) {
    if (value < 10) {
        return "0" + value;
    } else {
        return "" + value;
    }
}

/**
 * create a new text file.
 * if a file already exists, this program will exit.
 */
function createFile(filePath) {
    if (fso.fileExists(filePath)) {
        // abort
        echo("-- ABORT --\n\n"
             + "the following file does already exist.\n"
             + filePath);
        destroyObjects();
        WScript.quit();
    }

    var file = fso.createTextFile(filePath, overwrite = false);
    file.close();

    return fso.getFile(filePath);
}

/**
 * append new text into a file.
 * if the file does not exist, this program will exit.
 */
function appendToFile(filePath, text) {
    if (!fso.fileExists(filePath)) {
        // abort
        echo("ABORT\n\n"
             + "the following file does NOT already exist.\n"
             + filePath);
        destroyObjects();
        WScript.quit();
    }

    // iomode
    var ForReading = 1;
    var ForWriting = 2;
    var ForAppending = 8;

    // format
    var TristateTrue = -1; // Unicode
    var TristateFalse = 0; // ASCII
    var TristateUseDefault = -2; // System Default

    var file = fso.openTextFile(filePath, ForAppending, create = false, TristateUseDefault);
    file.writeLine(text);

    file.close();
}

/**
 * release several ActiveXObject(s).
 * this method should be called only when this program exits.
 */
function destroyObjects() {
    sho = null;
    fso = null;
}



//////////////////////////////////////////
//                                      //
//    JScript main procedure,           //
//    which is called from .bat file.   //
//                                      //
//////////////////////////////////////////

// global variables
var sho, fso;
var outputFileForOK, outputFileForNG, outputFileForWWW;

function main() {
    // initialize several ActiveXObject(s).
    sho = new ActiveXObject("WScript.Shell");
    fso = new ActiveXObject("Scripting.FileSystemObject");

    // get current directory as the target directory.
    var currentDirectory = sho.currentDirectory;

    // prepare specific files where the result will be written.
    var now = new Date();
    var dateString = now.getFullYear()
        + format2d(now.getMonth()+1)
        + format2d(now.getDate())
        + "-"
        + format2d(now.getHours())
        + format2d(now.getMinutes())
        + format2d(now.getSeconds());

    var baseName = fso.buildPath(sho.specialFolders("Desktop"),
                                 dateString + "_" + fso.getFolder(currentDirectory).name);

    outputFileForOK = createFile(baseName + "_ok.txt");
    outputFileForNG = createFile(baseName + "_ng.txt");
    outputFileForWWW = createFile(baseName + "_www.txt");

    // display basic information.
    stdout();
    stdout("********************************************************");
    stdout("  [local alive links] - OK cases will be output in this file:");
    stdout("    " + outputFileForOK.path);
    stdout()
    stdout("  [local dead links] - NG cases will be output in this file:");
    stdout("    " + outputFileForNG.path);
    stdout();
    stdout("  [links to \"The Internet\"] - Other cases will be output in this file:");
    stdout("    " + outputFileForWWW.path);
    stdout("********************************************************");
    stdout();

    // get all files recursively.
    var filePaths = enumFiles(currentDirectory);

    // check DEAD-LINKs per each file.
    for (var i in filePaths) {
        var filePath = filePaths[i];
        var msg = "--- " + filePath + " ---";

        stdout(msg);
        appendToFile(outputFileForOK.path, msg);
        appendToFile(outputFileForNG.path, msg);
        appendToFile(outputFileForWWW.path, msg);

        checkAllLinks(filePath);

        stdout();
        appendToFile(outputFileForOK.path, "");
        appendToFile(outputFileForNG.path, "");
        appendToFile(outputFileForWWW.path, "");
    }

    // release object.
    destroyObjects();

    // all done.
    stdout();
    stdout("*** all the tasks have been done !! ***")
    stdout("*** please follow the prompt message to quit this program. ***");
    stdout();
}

// execute the main function.
main();
