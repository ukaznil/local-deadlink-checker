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
    WScript.StdOut.writeLine(str);
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
    var document = new ActiveXObject("htmlfile");
    document.write(readFile(filePath));

    var wholeDocument = document.getElementsByTagName("html")[0];

    function checkLinksByTag(tagName, attrName) {
        var elements = document.getElementsByTagName(tagName);
        for (var i = 0, len = elements.length; i < len; i++) {
            var element = elements[i];

            // line number todo:
            var lineIndex = 0;

            var regex = new RegExp(attrName + "=\"([^\"]*)");
            var link = element.outerHTML.match(regex)[1];

            var status = getLinkStatus(link);
            var msg = (status === "NG" ? "  !!" : "    ")
                + "[" + status + "] Line " + lineIndex + ": " + element.outerHTML;

            // output to stdout.
            stdout(msg);

            // output to corresponding files.
            switch(status) {
            case "OK":
                appendToFile(outputFileForOK, msg);
                break;
            case "NG":
                appendToFile(outputFileForNG, msg);
                break;
            case "--":
                appendToFile(outputFileForWWW, msg);
                break;
            default:
                echo("Unknown error.");
                destroyObjects();
                WScript.quit();
                break;
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
function getLinkStatus(link) {
    if (link.startsWith("http://") ||
        link.startsWith("https://") ||
        link.startsWith("//")) {
        var status = "--";
    } else {
        var status = fso.fileExists(link) ? "OK" : "NG";
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
             + getLongPath(filePath));
        destroyObjects();
        WScript.quit();
    }

    var file = fso.createTextFile(filePath, overwrite = false);
    file.close();
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
             + getLongPath(filePath));
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
 * get file path in short format.
 */
function getShortPath(filePath) {
    var file = fso.getFile(filePath);
    return file.shortPath;
}

/**
 * get file path in long format.
 */
function getLongPath(filePath) {
    var file = fso.getFile(filePath);
    return file.path;
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

// initialize several ActiveXObject(s).
var sho = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");

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

var baseName = sho.specialFolders("Desktop") + "/"
    + dateString + "_" + fso.getFolder(currentDirectory).name;

var outputFileForWWW = baseName + "_www.txt";
createFile(outputFileForWWW);
var outputFileForOK = baseName + "_ok.txt";
createFile(outputFileForOK);
var outputFileForNG = baseName + "_ng.txt";
createFile(outputFileForNG);

// display basic information.
stdout();
stdout("********************************************************");
stdout("  [local alive links] - OK cases will be output in this file:");
stdout("    " + getLongPath(outputFileForOK));
stdout()
stdout("  [local dead links] - NG cases will be output in this file:");
stdout("    " + getLongPath(outputFileForNG));
stdout();
stdout("  [links to \"The Internet\"] - Other cases will be output in this file:");
stdout("    " + getLongPath(outputFileForWWW));
stdout("********************************************************");
stdout();

// get all files recursively.
var filePaths = enumFiles(currentDirectory);

// check DEAD-LINKs per each file.
for (var i in filePaths) {
    var filePath = filePaths[i];
    var msg = "--- " + filePath + " ---";

    stdout(msg);
    appendToFile(outputFileForOK, msg);
    appendToFile(outputFileForNG, msg);
    appendToFile(outputFileForWWW, msg);

    checkAllLinks(filePath);

    stdout();
    appendToFile(outputFileForOK, "");
    appendToFile(outputFileForNG, "");
    appendToFile(outputFileForWWW, "");
}

// release object.
destroyObjects();

// all done.
stdout();
stdout("*** all the tasks have been done !! ***")
stdout("*** please follow the prompt message to quit this program. ***");
stdout();
