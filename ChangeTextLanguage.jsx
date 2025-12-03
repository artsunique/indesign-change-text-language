/**
 * Text Language Converter für Adobe InDesign
 * Ändert die Sprache aller Textbereiche im Dokument
 */

#target indesign

(function() {
    // Hilfsfunktion: indexOf für Arrays (ExtendScript hat das nicht)
    function arrayIndexOf(arr, item) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === item) {
                return i;
            }
        }
        return -1;
    }

    // Prüfen ob ein Dokument geöffnet ist
    if (app.documents.length === 0) {
        alert("Bitte öffne zuerst ein Dokument.", "Kein Dokument");
        return;
    }

    var doc = app.activeDocument;

    // Verfügbare Sprachen sammeln
    var languages = app.languagesWithVendors;
    var languageNames = [];
    
    for (var i = 0; i < languages.length; i++) {
        languageNames.push(languages[i].name);
    }
    
    // Alphabetisch sortieren
    languageNames.sort();

    // Häufig verwendete Sprachen nach vorne
    var commonLanguages = [
        "German: 2006 Reform",
        "German: Swiss 2006 Reform", 
        "German: Austrian 2006 Reform",
        "English: USA",
        "English: UK",
        "Spanish: Castilian",
        "French",
        "Italian",
        "Portuguese",
        "Dutch",
        "Polish",
        "Czech",
        "Hungarian",
        "Romanian",
        "Turkish",
        "Russian",
        "Chinese",
        "Japanese",
        "Korean"
    ];
    
    // Gefundene häufige Sprachen nach vorne sortieren
    var sortedLanguages = [];
    for (var i = 0; i < commonLanguages.length; i++) {
        for (var j = 0; j < languageNames.length; j++) {
            if (languageNames[j] === commonLanguages[i]) {
                sortedLanguages.push(languageNames[j]);
                break;
            }
        }
    }
    
    // Trennlinie und Rest
    sortedLanguages.push("─────────────────────");
    for (var i = 0; i < languageNames.length; i++) {
        if (arrayIndexOf(sortedLanguages, languageNames[i]) === -1) {
            sortedLanguages.push(languageNames[i]);
        }
    }

    // Dialog erstellen
    var dialog = new Window("dialog", "Sprache ändern");
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];

    // Info-Text
    dialog.add("statictext", undefined, "Zielsprache für alle Textbereiche:");

    // Dropdown für Sprachauswahl
    var languageDropdown = dialog.add("dropdownlist", undefined, sortedLanguages);
    languageDropdown.preferredSize.width = 300;
    languageDropdown.selection = 0;

    // Optionen
    var optionsPanel = dialog.add("panel", undefined, "Optionen");
    optionsPanel.alignChildren = ["left", "top"];
    optionsPanel.margins = 15;
    
    var includeOverrides = optionsPanel.add("checkbox", undefined, "Auch lokale Formatierungen ändern");
    includeOverrides.value = true;
    
    var includeTables = optionsPanel.add("checkbox", undefined, "Tabellen einschließen");
    includeTables.value = true;

    // Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = ["right", "top"];
    var cancelBtn = buttonGroup.add("button", undefined, "Abbrechen", { name: "cancel" });
    var okBtn = buttonGroup.add("button", undefined, "Ändern", { name: "ok" });

    // Dialog anzeigen
    if (dialog.show() === 1) {
        var selectedName = languageDropdown.selection.text;
        
        // Trennlinie ignorieren
        if (selectedName.indexOf("───") !== -1) {
            alert("Bitte wähle eine gültige Sprache.", "Ungültige Auswahl");
            return;
        }

        // Sprache finden
        var targetLanguage = null;
        for (var i = 0; i < languages.length; i++) {
            if (languages[i].name === selectedName) {
                targetLanguage = languages[i];
                break;
            }
        }

        if (!targetLanguage) {
            alert("Sprache nicht gefunden.", "Fehler");
            return;
        }

        var changedStories = 0;
        var changedTables = 0;

        // Alle Stories durchgehen
        for (var i = 0; i < doc.stories.length; i++) {
            var story = doc.stories[i];
            
            try {
                // Gesamte Story auf neue Sprache setzen
                story.appliedLanguage = targetLanguage;
                changedStories++;

                // Lokale Formatierungen auch ändern
                if (includeOverrides.value) {
                    for (var j = 0; j < story.paragraphs.length; j++) {
                        var para = story.paragraphs[j];
                        para.appliedLanguage = targetLanguage;
                        
                        for (var k = 0; k < para.characters.length; k++) {
                            para.characters[k].appliedLanguage = targetLanguage;
                        }
                    }
                }

                // Tabellen
                if (includeTables.value && story.tables.length > 0) {
                    for (var t = 0; t < story.tables.length; t++) {
                        var table = story.tables[t];
                        for (var c = 0; c < table.cells.length; c++) {
                            table.cells[c].texts[0].appliedLanguage = targetLanguage;
                        }
                        changedTables++;
                    }
                }
            } catch (e) {
                // Manche Stories können nicht geändert werden (z.B. gesperrt)
            }
        }

        // Ergebnis
        var message = "Sprache geändert zu:\n" + selectedName + "\n\n";
        message += "Textbereiche: " + changedStories + "\n";
        if (includeTables.value) {
            message += "Tabellen: " + changedTables;
        }
        
        alert(message, "Fertig");
    }

})();
