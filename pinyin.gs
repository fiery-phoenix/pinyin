function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Converter', 'showSidebar')
        .addToUi();
}

function onInstall(e) {
    onOpen(e);
}

function showSidebar() {
    var ui = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('Pinyin')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    DocumentApp.getUi().showSidebar(ui);
}

function getSelectedText() {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (selection) {
        var text = [];
        var elements = selection.getSelectedElements();
        for (var i = 0; i < elements.length; i++) {
            if (elements[i].isPartial()) {
                var element = elements[i].getElement().asText();
                var startIndex = elements[i].getStartOffset();
                var endIndex = elements[i].getEndOffsetInclusive();

                text.push(element.getText().substring(startIndex, endIndex + 1));
            } else {
                var element = elements[i].getElement();
                // Only translate elements that can be edited as text; skip images and
                // other non-text elements.
                if (element.editAsText) {
                    var elementText = element.asText().getText();
                    // This check is necessary to exclude images, which return a blank
                    // text element.
                    if (elementText != '') {
                        text.push(elementText);
                    }
                }
            }
        }
        if (text.length == 0) {
            throw 'Please select some text.';
        }
        return text;
    } else {
        throw 'Please select some text.';
    }
}

function runConversion() {
    var text = getSelectedText();
    var resultText = "";
    for (var i = 0, len = text.length; i < len; i++) {
        resultText += pinyin.convertLine(text[i]) + '\n';
    }

    return resultText.trim();
}

var pinyin = (function () {
    var initials = "[bpmfdtnlgkhjqxrzcsyw]|zh|ch|sh";
    var finals = ["[aoeiuü]", "ai|ao|an|ang", "ou|ong", "ei|en|eng|er", "ia|ie|iao|iu|in|ian|iang|ing|iong",
        "ua|uo|uai|ui|uan|un|uang|ueng", "üe, üan, ün"];
    var tones = "[1234]";

    var pinyinMatcher = new RegExp("^((er|((" + initials + ")(" + finals.join("|") + ")r?))(" + tones + "*))+$", "i");
    var numericPinyinMatcher = new RegExp("(er|((" + initials + ")(" + finals.join("|") + ")r?))(" + tones + "+)", "ig");

    var vowelsToVowelsWithTones = {
        a: {1: "ā", 2: "á", 3: "ǎ", 4: "à"},
        o: {1: "ō", 2: "ó", 3: "ŏ", 4: "ò"},
        E: {1: "Ē", 2: "É", 3: "Ě", 4: "È"},
        e: {1: "ē", 2: "é", 3: "ě", 4: "è"},
        ui: {1: "uī", 2: "uí", 3: "uǐ", 4: "uì"},
        iu: {1: "iū", 2: "iú", 3: "iǔ", 4: "iù"},
        i: {1: "ī", 2: "í", 3: "ǐ", 4: "ì"},
        u: {1: "ū", 2: "ú", 3: "ǔ", 4: "ù"},
        ü: {1: "ǖ", 2: "ǘ", 3: "ǚ", 4: "ǜ"}
    };

    var vowelsRegularExpressions = [
        ["a", /a(.*)\d/i, "$1"],
        ["o", /o(.*)\d/i, "$1"],
        ["E", /E(r)\d/, "$1"],
        ["e", /e(.*)\d/i, "$1"],
        ["ui", /ui\d/i, ""],
        ["iu", /iu\d/i, ""],
        ["i", /i(.*)\d/i, "$1"],
        ["u", /u(.*)\d/i, "$1"],
        ["ü", /ü(.*)\d/i, "$1"]
    ];

    return {
        isNumericPinyin: function (word) {
            return pinyinMatcher.test(word) && /[1234]/.test(word);
        },

        convertSyllable: function (syllable) {
            for (var i = 0; i < vowelsRegularExpressions.length; i++) {
                var vowelRegularExpression = vowelsRegularExpressions[i];
                var convertedSyllable = syllable.replace(vowelRegularExpression[1],
                    vowelsToVowelsWithTones[vowelRegularExpression[0]][syllable.slice(-1)] + vowelRegularExpression[2]);
                if (convertedSyllable != syllable) {
                    return convertedSyllable;
                }
            }
            return syllable;
        },

        convertWord: function (word) {
            return word.replace(numericPinyinMatcher, pinyin.convertSyllable);
        },

        convertLine: function (line) {
            return line.replace(/\b\w+\b/g, function (match, capture) {
                return pinyin.isNumericPinyin(match) ? pinyin.convertWord(match) : match;
            });
        }
    }
}());

function insertText(newText) {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (selection) {
        var replaced = false;
        var elements = selection.getSelectedElements();
        if (elements.length == 1 &&
            elements[0].getElement().getType() ==
            DocumentApp.ElementType.INLINE_IMAGE) {
            throw "Can't insert text into an image.";
        }
        for (var i = 0; i < elements.length; i++) {
            if (elements[i].isPartial()) {
                var element = elements[i].getElement().asText();
                var startIndex = elements[i].getStartOffset();
                var endIndex = elements[i].getEndOffsetInclusive();

                var remainingText = element.getText().substring(endIndex + 1);
                element.deleteText(startIndex, endIndex);
                if (!replaced) {
                    element.insertText(startIndex, newText);
                    replaced = true;
                } else {
                    // This block handles a selection that ends with a partial element. We
                    // want to copy this partial text to the previous element so we don't
                    // have a line-break before the last partial.
                    var parent = element.getParent();
                    parent.getPreviousSibling().asText().appendText(remainingText);
                    // We cannot remove the last paragraph of a doc. If this is the case,
                    // just remove the text within the last paragraph instead.
                    if (parent.getNextSibling()) {
                        parent.removeFromParent();
                    } else {
                        element.removeFromParent();
                    }
                }
            } else {
                var element = elements[i].getElement();
                if (!replaced && element.editAsText) {
                    // Only translate elements that can be edited as text, removing other
                    // elements.
                    element.clear();
                    element.asText().setText(newText);
                    replaced = true;
                } else {
                    // We cannot remove the last paragraph of a doc. If this is the case,
                    // just clear the element.
                    if (element.getNextSibling()) {
                        element.removeFromParent();
                    } else {
                        element.clear();
                    }
                }
            }
        }
    } else {
        var cursor = DocumentApp.getActiveDocument().getCursor();
        var surroundingText = cursor.getSurroundingText().getText();
        var surroundingTextOffset = cursor.getSurroundingTextOffset();

        // If the cursor follows or preceds a non-space character, insert a space
        // between the character and the translation. Otherwise, just insert the
        // translation.
        if (surroundingTextOffset > 0) {
            if (surroundingText.charAt(surroundingTextOffset - 1) != ' ') {
                newText = ' ' + newText;
            }
        }
        if (surroundingTextOffset < surroundingText.length) {
            if (surroundingText.charAt(surroundingTextOffset) != ' ') {
                newText += ' ';
            }
        }
        cursor.insertText(newText);
    }
}
