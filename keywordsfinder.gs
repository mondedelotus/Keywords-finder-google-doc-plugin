// --Globals--
// Current File
var ui = DocumentApp.getUi();
var doc = DocumentApp.getActiveDocument();

function onOpen(e) {
    DocumentApp.getUi()
        .createAddonMenu()
        .addItem("Start", "showSidebar")
        .addToUi();
}

function onInstall(e) {
    onOpen(e);
}

function showSidebar() {
    var html = HtmlService.createTemplateFromFile("main")
        .evaluate()
        .setTitle("KeywordsFinder")
        .setWidth(400);
    ui.showSidebar(html);
}

// Creates an import or include function so files can be added
// inside the main index.
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAllCounts(
    keywords,
    phrases,
    ignoreCase,
    ignoreWords,
    ignorePhrases,
    deleteEmptyRows,
    wordsHighlight,
    phrasesHighlight
) {
    if (!keywords && !phrases) return {};
    keywords = keywords
        .replace(/(\s|\n|\r)/g, ",")
        .split(",")
        .map(function (item) {
            return item.trim();
        })
        .filter(Boolean);

    phrases = phrases
        .replace(/(\n|\r)/g, ",")
        .split(",")
        .map(function (item) {
            return item.trim();
        })
        .filter(Boolean)
        .sort(function (a, b) {
            return a.localeCompare(b);
        });

    var phrasesCount = {};
    phrases.forEach(function (x) {
        phrasesCount[x] = (phrasesCount[x] || 0) + 1;
    });

    if (deleteEmptyRows) {
        removeExtraSpaces();
        removeEmptyRows();
    }

    var phrases_summary = findWords(
        phrases,
        null,
        ignoreCase,
        null,
        wordsHighlight,
        phrasesHighlight
    );

    var words_summary = findWords(
        keywords,
        phrases_summary,
        ignoreCase,
        ignoreWords,
        wordsHighlight,
        phrasesHighlight
    );
    Logger.log("summary before", words_summary, phrases_summary);
    if (ignorePhrases) {
        phrases_summary.found = extractPhrasesDuplicates(
            phrases_summary.found,
            ignoreCase
        );
    }

    phrases_summary.found = addDuplicatesCount(
        phrases_summary.found,
        phrasesCount
    );
    Logger.log("summary", words_summary, phrases_summary);

    var result = formatOutput(words_summary, phrases_summary);
    result.hasNotFound =
        words_summary.hasNotFound || phrases_summary.hasNotFound;
    return result;
}

function removeArrayDuplicates(arr, ignoreCase) {
    if (ignoreCase) {
        arr = arr.map(function (item) {
            return item.toLowerCase();
        });
    }

    arr.filter(function (item, pos, self) {
        return self.indexOf(item) === pos;
    });
}

function addDuplicatesCount(phrases, count) {
    var result = {};
    for (var attrname in phrases) {
        var key = attrname + "[" + count[attrname] + "]";
        result[key] = phrases[attrname];
    }
    return result;
}

function objToString(obj, onlyWord) {
    var result = "";
    for (var attrname in obj) {
        result += attrname + (onlyWord ? "" : ": " + obj[attrname]) + "\n";
    }
    return result;
}

function formatOutput(words, phrases) {
    var found = "___Found Phrases___\n";
    found += objToString(phrases.found);
    found += "\n___Found Words___\n";
    found += objToString(words.found);

    var notFound = "___Not Found Phrases___\n";
    notFound += objToString(phrases.notFound, true);
    notFound += "\n___Not Found Words___\n";
    notFound += objToString(words.notFound, true);

    return { found: found, notFound: notFound };
}

//phrases_summary parameter indicates if we search for words or phrases
function findWords(
    keys,
    phrases_summary,
    ignoreCase,
    ignoreWords,
    wordsHighlight,
    phrasesHighlight
) {
    var body = doc.getBody();
    var keysMap = { found: {}, notFound: {}, hasNotFound: false }; // object for keys with quantity
    // For every word in keys:
    Logger.log(ignoreCase);
    for (var w = 0; w < keys.length; ++w) {
        if (ignoreCase) {
            var regex = "(^|\\W)(?i)\\Q" + keys[w] + "\\E(^|$|\\W)";
        } else {
            var regex = "(^|\\W)\\Q" + keys[w] + "\\E(^|$|\\W)";
        }

        Logger.log(regex);

        var foundElement = body.findText(regex);
        var count = 0;

        while (foundElement != null) {
            // Get the text object from the element
            var foundText = foundElement.getElement().asText();

            // Where in the Element is the found text?
            var start = foundElement.getStartOffset();
            var end = foundElement.getEndOffsetInclusive();

            count++;

            // Change the background color
            if (phrases_summary) {
                if (wordsHighlight) {
                    if (
                        foundText.getBackgroundColor(start) !==
                            phrasesHighlight ||
                        phrasesHighlight === null
                    ) {
                        foundText.setBackgroundColor(
                            start,
                            end,
                            wordsHighlight
                        );
                    }
                }
            } else {
                if (phrasesHighlight) {
                    foundText.setBackgroundColor(start, end, phrasesHighlight);
                }
            }

            // Find the next match
            foundElement = body.findText(regex, foundElement);
        }

        if (phrases_summary && ignoreWords) {
            count -= getFoundPhrasesCount(phrases_summary, keys[w], ignoreCase);
        }

        if (count > 0) {
            keysMap.found[keys[w]] = count;
        } else {
            keysMap.notFound[keys[w]] = count;
            keysMap.hasNotFound = true;
        }
    }

    return keysMap;
}

function getFoundPhrasesCount(phrases_summary, word, ignoreCase) {
    var count = 0;

    var regex = new RegExp(
        "(^|\\s)" +
            word.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&") +
            "(^|$|\\s)",
        ignoreCase ? "gi" : "g"
    );
    for (var attrname in phrases_summary.found) {
        Logger.log("count:", attrname, (attrname.match(regex) || []).length);
        if (attrname.search(regex) !== -1) {
            count += phrases_summary.found[attrname];
        }
    }
    return count;
}

function extractPhrasesDuplicates(foundPhrases, ignoreCase) {
    var regex,
        result = {};

    for (var curPhrase in foundPhrases) {
        var count = foundPhrases[curPhrase];
        regex = new RegExp(
            "(^|\\s)" +
                curPhrase.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&") +
                "(^|$|\\s)",
            ignoreCase ? "gi" : "g"
        );

        for (var phrase in foundPhrases) {
            if (
                phrase.search(regex) !== -1 &&
                (ignoreCase
                    ? phrase.toLowerCase() !== curPhrase.toLowerCase()
                    : phrase !== curPhrase)
            ) {
                count -= foundPhrases[phrase];
            }
        }
        result[curPhrase] = count;
    }
    return result;
}

function getAllSettings() {
    return PropertiesService.getUserProperties().getProperties();
}

function saveSettings(key, value) {
    PropertiesService.getUserProperties().setProperty(key, value);
}

function deleteSettings(key) {
    PropertiesService.getUserProperties().deleteProperty(key);
}

function deleteAllSettings() {
    PropertiesService.getUserProperties().deleteAllProperties();
}

function removeEmptyRows() {
    var pars = DocumentApp.getActiveDocument().getBody().getParagraphs();
    // for each paragraph in the active document...
    pars.forEach(function (e) {
        try {
            // does the paragraph contain an image or a horizontal rule?
            // (you may want to add other element types to this check)
            no_img =
                e.findElement(DocumentApp.ElementType.INLINE_IMAGE) === null;
            no_rul =
                e.findElement(DocumentApp.ElementType.HORIZONTAL_RULE) === null;
            // proceed if it only has text
            if (no_img && no_rul) {
                // clean up paragraphs that only contain whitespace
                e.replaceText("^\\s+$", "");
                // remove blank paragraphs
                if (e.getText() === "") {
                    e.removeFromParent();
                }
            }
        } catch (e) {}
    });
}

function removeExtraSpaces() {
    DocumentApp.getActiveDocument().getBody().replaceText("\\s+", " ");
}

function setTextStyle(font, fontSize, pIndent) {
    var body = DocumentApp.getActiveDocument().getBody();
    // var textStyle = body.getHeadingAttributes(
    //   DocumentApp.ParagraphHeading["NORMAL"]
    // );
    var textStyle = {};
    Logger.log(textStyle);
    textStyle[DocumentApp.Attribute.FONT_FAMILY] = font;
    textStyle[DocumentApp.Attribute.FONT_SIZE] = fontSize;
    textStyle[DocumentApp.Attribute.SPACING_AFTER] = pIndent;
    // body.setHeadingAttributes(DocumentApp.ParagraphHeading["NORMAL"], textStyle);

    body.getParagraphs().forEach(function (p) {
        if (p.getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
            Logger.log("normal");
            p.setAttributes(textStyle);
            // p.setSpacingAfter(parseFloat(pIndent));
        }
    });
}

function setHeadingsStyle(headingsType, font, fontSize, isBold, isItalic) {
    var body = DocumentApp.getActiveDocument().getBody();

    // const headings = [
    //     DocumentApp.ParagraphHeading['HEADING1'],
    //     DocumentApp.ParagraphHeading['HEADING2'],
    //     DocumentApp.ParagraphHeading['HEADING3'],
    //     DocumentApp.ParagraphHeading['HEADING4'],
    //     DocumentApp.ParagraphHeading['HEADING5'],
    //     DocumentApp.ParagraphHeading['HEADING6']
    // ];

    var textStyle = {};
    Logger.log(textStyle);
    textStyle[DocumentApp.Attribute.FONT_FAMILY] = font;
    textStyle[DocumentApp.Attribute.FONT_SIZE] = fontSize;
    textStyle[DocumentApp.Attribute.BOLD] = isBold;
    textStyle[DocumentApp.Attribute.ITALIC] = isItalic;

    body.getParagraphs().forEach(function (p) {
        // if (headings.indexOf(p.getHeading()) !== -1) {
        //     Logger.log('header');
        //     p.setAttributes(textStyle);
        // }
        Logger.log(p.getHeading());
        if (p.getHeading() === DocumentApp.ParagraphHeading[headingsType]) {
            p.setAttributes(textStyle);
        }
    });
}
