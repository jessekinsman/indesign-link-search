// This script uses the links collection to search links
// It might be faster than the alternate of looping through all pageItems in document

// TO DO
// Test performance agains looping through pageItems
// Add results into staticLabels that can be clicked to view the link
// Use click event on labels
// Pass in function to open the document, then show link using link.show()
/*global UserInteractionLevels:false, File:false, Window:false, SaveOptions:false, alert: false, FitOptions:false, app:false */

/* jshint ignore:start */
#targetengine "session";
/* jshint ignore:end */


function findLinks(term) {
    var indexObj = {},
        myDoc, book, matches = [],
        loopDoc, myBook, myDocs, myFolio;
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    if (app.books.length > 0) {
        myBook = app.activeBook;
        book = myBook.bookContents;
    } else {
        book = app.activeDocument;
    }
    loopDoc = function (doc) {
        var pageItem, linkName, pageNum = "on pasteboard",
            linkPath, linked, linkArr, par, parPage, docPath, linkId;
        linkArr = doc.links;
        //alert("link array length " + linkArr.length);
        for (var idx = 0; idx < linkArr.length; idx++) {
            pageItem = linkArr.item(idx);
            linkName = pageItem.name;
            if (linkName.toLowerCase().search(term.toLowerCase()) > -1) {
                //alert("link name: " + linkName);
                //pageItem.show();
                linkId = pageItem.id;
                par = pageItem.parent;
                //alert("parent type " + par.constructor.name);
                if (par.constructor.name === "Story") {
                    //alert("page items length " + par.pageItems.length);
                    for (var txtl = 0; txtl < par.textContainers.length; txtl++) {
                        parPage = par.textContainers[txtl].parentPage;
                        if (parPage !== null) {
                            txtl = par.textContainers.length;
                            pageNum = "on page " + parPage.name;
                        }
                    }
                } else {
                    parPage = par.parentPage;
                    if (parPage !== null && parPage.name !== null) {
                        pageNum = "on page " + parPage.name;
                    } else {
                        pageNum = "on Pasteboard";
                    }
                }

                //alert("we have a match " + linkName + " on page " + pageNum);
                matches.push({
                    "text": linkName + ", " + pageNum,
                    "docPath": doc.fullName,
                    "id": linkId
                });
            }
        }
        closeDoc(doc);
    };
    if (book.reflect.name == "BookContents") {
        for (var docs = 0; docs < book.length; docs++) {
            myDoc = book.item(docs);
            // uncomment below to work through book
            myDoc = app.open(File(myDoc.fullName), false);
            loopDoc(myDoc);
        }
    } else {
        loopDoc(book);
    }
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    showMissing(matches);
}

function closeDoc(doc) {
    if (doc.visible === false) {
        doc.close(SaveOptions.yes);
    }
}

function showMissing(missing, linked) {
    //alert("missing: " + missing.length + " linked: " + linked.length);
    showInfoDialog(missing, linked);
}

function showInfoDialog(matches) {
    var missingForm, missingStr = "",
        linkedStr = "",
        result, mText, lText, header,
        bgroup, mGroup, lGroup, resultLabel = [],
        resultButton = [],
        myDialog, rGroup;
    if (matches.length > 0) {
        myDialog = Window.find("palette", "Search Results");
        if (myDialog !== null) {
            myDialog.close();
            myDialog = null;
            delete myDialog;
            myDialog = new Window("palette", "Search Results", undefined, {
                closeButton: true
            });
        } else {
            myDialog = new Window("palette", "Search Results", undefined, {
                closeButton: true
            });
        }
        header = myDialog.add("staticText", undefined, "Search Results (click to see link)");
        header.graphics.font = "Tahoma-Bold:14";
        for (var i = 0; i < matches.length; i++) {
            lGroup = myDialog.add("panel");
            lGroup.orientation = "row";
            resultLabel[i] = lGroup.add("staticText", undefined, matches[i].text);
            resultButton[i] = lGroup.add("button", undefined, "View");
            resultButton[i].properties = matches[i];
            resultButton[i].addEventListener("click", function (e) {
                var linkArr, sLink, myDoc;
                var props = e.target.properties;
                app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
                myDoc = app.open(File(props.docPath), true);
                linkArr = myDoc.links;
                linkArr.itemByID(props.id).show();
                app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
            });
        }
        myDialog.show();
    } else {
        alert("0 Matches");
    }
}

function searchDialog(missing, linked) {
    var myDialog = new Window("dialog", "Search Links", undefined, {
            closeButton: true
        }),
        linkedStr = "",
        result, mText, lText,
        bGroup, mGroup, lGroup, term;
    lGroup = myDialog.add("group");
    lGroup.orientation = "row";
    lText = lGroup.add("editText", [0, 0, 300, 25], "", {
        multiline: false
    });
    lText.text = "Search Term";
    lText.active = true;
    bGroup = myDialog.add("group");
    bGroup.add("button", undefined, "OK");
    bGroup.add("button", undefined, "Cancel");
    result = myDialog.show();
    if (result == 1) {
        term = lText.text;
        if (term == "Search Term" || term === "") {
            alert("Enter a search term");
        } else {
            findLinks(term);
        }
    }
}

function linkFile(file, path, frame) {
    var testFile, adPath, extAr = ["jpg", "pdf", "eps", "ai", "tif"],
        exist = false;
    for (var e = 0; e < extAr.length; e++) {
        testFile = new File(path + file + "." + extAr[e]);
        if (testFile.exists) {
            exist = true;
            frame.place(testFile, false);
            frame.fit(FitOptions.CONTENT_TO_FRAME);
            e = extAr.length;
        }
    }
    return exist;
}

function getIssueNum(doc) {
    var textVar = doc.textVariables,
        issueNum = null;
    for (var txtL = 0; txtL < textVar.length; txtL++) {
        if (textVar.item(txtL).name === "Issue") {
            issueNum = textVar.item(txtL).variableOptions.contents;
        }
    }
    return issueNum;
}

function trim(str) {
    return str.replace(/^\s+|\s+$/g, '');
}

function inspectObject(targ) {
    for (var key in targ) {
        if (targ.hasOwnProperty(key)) {
            alert(key + " -> " + targ[key]);
        }
    }
}

searchDialog();