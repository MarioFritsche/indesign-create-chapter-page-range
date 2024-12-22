// Hauptfunktion zum Erstellen der Kapitelübersicht oder des klassischen Inhaltsverzeichnisses
function createChapterPageRange() {
  var doc = app.activeDocument;

  // JSON-Polyfill für ExtendScript
  var JSON = {
    stringify: function (obj) {
      var str = "{";
      for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
          var value = obj[key];
          str += '"' + key + '":';
          if (typeof value === "string") {
            str += '"' + value + '"';
          } else if (typeof value === "object") {
            str += JSON.stringify(value);
          } else {
            str += value;
          }
          str += ",";
        }
      }
      return str.replace(/,$/, "") + "}";
    },

    parse: function (str) {
      return eval("(" + str + ")");
    },
  };

  // Überprüfen, ob ein Dokument geöffnet ist
  if (!doc) {
    alert("Bitte ein InDesign-Dokument öffnen.");
    return;
  }

  // Überprüfen, ob ein Textrahmen ausgewählt ist
  var selectedTextFrame =
    app.selection.length === 1 &&
    app.selection[0].constructor.name === "TextFrame"
      ? app.selection[0]
      : null;

  // Überprüfen, ob der ausgewählte Textrahmen leer ist
  var isTextFrameEmpty =
    selectedTextFrame && selectedTextFrame.contents.length === 0;

  // Textrahmen mit dem Label "ChapterOverviewFrame" suchen
  var existingTextFrames = []; // Initialisieren
  for (var i = 0; i < doc.textFrames.length; i++) {
    try {
      var labelData = JSON.parse(doc.textFrames[i].label);
      if (labelData && labelData.type === "ChapterOverviewFrame") {
        existingTextFrames.push(doc.textFrames[i]);
      }
    } catch (e) {
      continue;
    }
  }

  // Wenn weder ein leerer Textrahmen ausgewählt ist noch Textrahmen mit dem Label "ChapterOverviewFrame" existieren
  if (existingTextFrames.length === 0 && !isTextFrameEmpty) {
    alert(
      "Bitte erstellen Sie einen leeren Textrahmen oder verwenden Sie einen Textrahmen mit dem Label 'ChapterOverviewFrame'."
    );
    return;
  }

  // Dialog erstellen
  var dialog = new Window(
    "dialog",
    "Inhaltsverzeichnis und Kapitelübersicht erstellen"
  );
  dialog.preferredSize.width = 392;
  dialog.alignChildren = ["left", "top"];
  // Hauptgruppe zur linken Anordnung
  var mainGroup = dialog.add("group");
  mainGroup.orientation = "column";
  mainGroup.alignChildren = "left";

  // Panel für Formatwahl (Inhaltsverzeichnis oder Kapitelübersicht)
  var formatPanel = mainGroup.add("panel", undefined, "Format auswählen");
  formatPanel.alignChildren = "left";
  formatPanel.preferredSize.width = 350;
  var formatGroup = formatPanel.add("group");

  var radioButtonTableOfContents = formatGroup.add(
    "radiobutton",
    undefined,
    "Inhaltsverzeichnis"
  );
  var radioButtonChapterOverview = formatGroup.add(
    "radiobutton",
    undefined,
    "Kapitelübersicht"
  );

  // Standard Inhaltsverzeichnis
  radioButtonTableOfContents.value = true;

  // Neues Panel "Allgemein" mit den Feldern
  var generalPanel = mainGroup.add("panel", undefined, "Allgemein");
  generalPanel.alignChildren = "left";
  generalPanel.preferredSize.width = 350;

  generalPanel.add("statictext", undefined, "Titel");
  var titleInput = generalPanel.add(
    "edittext",
    undefined,
    "Kapitelübersicht mit Seitenbereichen:"
  );
  titleInput.characters = 30;

  generalPanel.add(
    "statictext",
    undefined,
    "Welches Absatzformat soll genutzt werden:"
  );
  var paragraphStyleDropdown = generalPanel.add(
    "dropdownlist",
    undefined,
    getAllParagraphStylesWithPath(doc)
  );
  paragraphStyleDropdown.selection = 0; // Standardmäßig das erste Element auswählen

  generalPanel.add("statictext", undefined, "Text vor der Seitenzahl:");
  var prefixInput = generalPanel.add("edittext", undefined, "Seite");
  prefixInput.characters = 15;

  generalPanel.add("statictext", undefined, "Text zwischen den Seitenzahlen:");
  var separatorInput = generalPanel.add("edittext", undefined, "-");
  separatorInput.characters = 15;
  separatorInput.enabled = false;

  // Panel für Optionen (Absatzformate Titel und Einträge)
  var optionsPanel = mainGroup.add("panel", undefined, "Optionen");
  optionsPanel.alignChildren = "left";
  optionsPanel.preferredSize.width = 350;
  optionsPanel.add("statictext", undefined, "Absatzformat Titel:");
  var titleParagraphStyleDropdown = optionsPanel.add(
    "dropdownlist",
    undefined,
    getAllParagraphStylesWithPath(doc)
  );
  titleParagraphStyleDropdown.selection = 0;

  optionsPanel.add("statictext", undefined, "Absatzformat Einträge:");
  var entryParagraphStyleDropdown = optionsPanel.add(
    "dropdownlist",
    undefined,
    getAllParagraphStylesWithPath(doc)
  );
  entryParagraphStyleDropdown.selection = 0;

  // Checkbox für "Vorhandenes Inhaltsverzeichnis aktualisieren"
  var updateCheckbox = optionsPanel.add(
    "checkbox",
    undefined,
    "Vorhandenes Inhaltsverzeichnis aktualisieren"
  );

  // Text für "Wählen Sie ein Inhaltsverzeichnis:"
  optionsPanel.add(
    "statictext",
    undefined,
    "Wählen Sie ein Inhaltsverzeichnis:"
  );
  // Dropdown für die Textrahmen-Auswahl (zu Beginn leer)
  var existingFramesDropdown = optionsPanel.add("dropdownlist", undefined, []);
  existingFramesDropdown.selection = 0; // Standardmäßig den ersten Textrahmen auswählen
  existingFramesDropdown.enabled = false; // Initial deaktiviert
  existingFramesDropdown.preferredSize.width = 150;

  // Funktion, um die erste Zeile eines Textrahmens zu extrahieren
  function getFirstLineOfTextFrame(frame) {
    if (!frame || !frame.contents) {
      return "Unbenannter Textrahmen";
    }
    var lines = frame.contents.split("\r"); // Nur die erste Zeile extrahieren
    return lines[0] || "Unbenannter Textrahmen";
  }

  // Textrahmen-Inhalte in das Dropdown laden
  var frameTitles = [];
  if (existingTextFrames.length > 0) {
    for (var i = 0; i < existingTextFrames.length; i++) {
      var title = getFirstLineOfTextFrame(existingTextFrames[i]);
      frameTitles.push(title);
      existingFramesDropdown.add("item", title);
    }
  } else {
    existingFramesDropdown.add("item", "Kein Inhaltsverzeichnis gefunden");
  }

  existingFramesDropdown.onChange = function () {
    var selectedFrame =
      existingTextFrames[existingFramesDropdown.selection.index];
    if (selectedFrame) {
      var labelData = JSON.parse(selectedFrame.label);
      // Den Wert von formatType aus dem Label lesen
      if (labelData.formatType === "chapterOverview") {
        radioButtonChapterOverview.value = true;
        separatorInput.enabled = true; // Feld für "Text zwischen den Seitenzahlen" aktivieren
      } else if (labelData.formatType === "classicTOC") {
        radioButtonTableOfContents.value = true;
        separatorInput.enabled = false; // Feld für "Text zwischen den Seitenzahlen" deaktivieren
      }
      // Titel, Prefix und Separator vorbefüllen
      titleInput.text = labelData.title || "";
      prefixInput.text = labelData.prefixText || "";
      separatorInput.text = labelData.separatorText || "";

      // Absatzformat für das Inhaltsverzeichnis aus dem Label holen und im Dropdown setzen
      var selectedStyle = labelData.paragraphStyle;
      var index = paragraphStyleDropdown.find(selectedStyle);

      // Überprüfen, ob der Stil im Dropdown vorhanden ist
      if (index !== -1) {
        // Stil im Dropdown auswählen
        paragraphStyleDropdown.selection = index;
      } else {
        // Wenn der Stil nicht gefunden wird, Fallback auf den ersten Eintrag im Dropdown
        alert(
          "Das Absatzformat '" +
            selectedStyle +
            "' wurde nicht gefunden. Es wird das erste Format ausgewählt."
        );
        paragraphStyleDropdown.selection = 0; // Setzt das erste Format als Fallback
      }

      // Vorab gesetzte Absatzformate für Titel und Einträge
      var titleParagraphStyle = labelData.titleParagraphStyle;
      var entryParagraphStyle = labelData.entryParagraphStyle;

      // Vorbelegung der Dropdowns für Titel und Einträge
      var titleStyleIndex =
        titleParagraphStyleDropdown.find(titleParagraphStyle);
      if (titleStyleIndex !== -1) {
        titleParagraphStyleDropdown.selection = titleStyleIndex;
      } else {
        titleParagraphStyleDropdown.selection = 0; // Fallback auf erstes Format
      }

      var entryStyleIndex =
        entryParagraphStyleDropdown.find(entryParagraphStyle);
      if (entryStyleIndex !== -1) {
        entryParagraphStyleDropdown.selection = entryStyleIndex;
      } else {
        entryParagraphStyleDropdown.selection = 0; // Fallback auf erstes Format
      }
    }
  };

  // Funktion, um den Index eines Absatzformats im Dropdown zu finden
  function findParagraphStyleIndex(dropdown, styleName) {
    for (var i = 0; i < dropdown.items.length; i++) {
      if (dropdown.items[i].text === styleName) {
        return i;
      }
    }
    return -1; // Rückgabe -1, wenn der Stil nicht gefunden wurde
  }

  // Checkbox-Änderung überwachen, um das Dropdown zu aktivieren
  updateCheckbox.onClick = function () {
    existingFramesDropdown.enabled = updateCheckbox.value; // Aktiviert das Dropdown nur wenn die Checkbox aktiviert ist
  };

  // Wenn kein Textrahmen mit dem Label "ChapterOverviewFrame" existiert, Update-Checkbox deaktivieren
  if (existingTextFrames.length === 0) {
    updateCheckbox.enabled = false;
  }

  // Event-Handler für den Radio-Button für Inhaltsverzeichnis
  radioButtonTableOfContents.onClick = function () {
    if (radioButtonTableOfContents.value) {
      separatorInput.enabled = false; // Deaktiviert das Textfeld für "Text zwischen den Seitenzahlen"
    }
  };

  // Event-Handler für den Radio-Button für Kapitelübersicht
  radioButtonChapterOverview.onClick = function () {
    if (radioButtonChapterOverview.value) {
      separatorInput.enabled = true; // Aktiviert das Textfeld für "Text zwischen den Seitenzahlen"
    }
  };

  // OK-Button
  var buttonGroup = dialog.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.alignChildren = ["right", "center"];
  buttonGroup.spacing = 10;
  buttonGroup.margins = 10;
  buttonGroup.alignment = ["right", "top"];
  var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });
  okButton.onClick = function () {
    dialog.close(1);
  };

  // Abbrechen-Button
  var cancelButton = buttonGroup.add("button", undefined, "Abbrechen", {
    name: "cancel",
  });

  // Dialog anzeigen
  if (dialog.show() !== 1) {
    return;
  }

  // Benutzeroptionen speichern
  var selectedStylePath = paragraphStyleDropdown.selection.text;
  var customTitle = titleInput.text;
  var prefixText = prefixInput.text;
  var separatorText = separatorInput.text;
  var updateExisting = updateCheckbox.value;

  var selectedFrameIndex = existingFramesDropdown.selection
    ? existingFramesDropdown.selection.index
    : -1;
  var selectedFrame =
    selectedFrameIndex >= 0 ? existingTextFrames[selectedFrameIndex] : null;

  // Absatzformat über den vollständigen Pfad finden
  var chapterStyle = findParagraphStyleByPath(doc, selectedStylePath);

  if (!chapterStyle || !chapterStyle.isValid) {
    alert(
      "Das Absatzformat '" +
        selectedStylePath +
        "' konnte nicht gefunden werden."
    );
    return;
  }

  // Wenn Absatzformat für Titel und Einträge ausgewählt wurden
  var titleParagraphStyle = titleParagraphStyleDropdown.selection
    ? titleParagraphStyleDropdown.selection.text
    : null;
  var entryParagraphStyle = entryParagraphStyleDropdown.selection
    ? entryParagraphStyleDropdown.selection.text
    : null;

  // Überprüfen, ob ein vorhandenes Inhaltsverzeichnis aktualisiert werden soll
  var existingTextFrame = null;

  if (updateExisting && selectedFrame) {
    existingTextFrame = selectedFrame; // Den ausgewählten Textrahmen verwenden
  }

  // Variablen für Kapitel und Seitenbereiche
  var chapters = [];
  var currentChapter = null;

  // Textrahmen durchlaufen
  for (var i = 0; i < doc.stories.length; i++) {
    var story = doc.stories[i];

    for (var j = 0; j < story.paragraphs.length; j++) {
      var paragraph = story.paragraphs[j];

      // Kapitelüberschrift gefunden
      if (paragraph.appliedParagraphStyle == chapterStyle) {
        var paragraphText = paragraph.contents
          .replace(/[\r\n]+/g, " ")
          .replace(/\s+/g, " ");
        if (!paragraphText) paragraphText = "Unbenanntes Kapitel";

        if (
          paragraph.parentTextFrames.length > 0 &&
          paragraph.parentTextFrames[0].parentPage
        ) {
          var currentPage = paragraph.parentTextFrames[0].parentPage.name;

          if (currentChapter) {
            currentChapter.endPage = currentPage;
            chapters.push(currentChapter);
          }

          currentChapter = {
            title: paragraphText,
            startPage: currentPage,
            endPage: null,
          };
        }
      }
    }
  }

  if (currentChapter) {
    currentChapter.endPage = doc.pages[doc.pages.length - 1].name;
    chapters.push(currentChapter);
  }

  // Überarbeiten des Abschnitts, der die Kapitelübersicht oder das Inhaltsverzeichnis erstellt
  var outputText = customTitle + "\r";

  // Kapitelübersicht erstellen
  if (radioButtonChapterOverview.value) {
    for (var k = 0; k < chapters.length; k++) {
      var chapter = chapters[k];
      outputText +=
        chapter.title +
        "\t" +
        prefixText +
        " " +
        chapter.startPage +
        " " +
        separatorText +
        " " +
        chapter.endPage +
        "\r";
    }
  }

  // Klassisches Inhaltsverzeichnis erstellen
  if (radioButtonTableOfContents.value) {
    for (var k = 0; k < chapters.length; k++) {
      var chapter = chapters[k];
      outputText +=
        chapter.title + "\t" + prefixText + " " + chapter.startPage + "\r";
    }
  }

  // Textrahmen einfügen und Label speichern
  if (!existingTextFrame) {
    if (selectedTextFrame) {
      // Text in den ausgewählten Textrahmen einfügen

      selectedTextFrame.contents = outputText;

      // Textabsätze formatieren
      var paragraphs = selectedTextFrame.paragraphs;
      for (var i = 0; i < paragraphs.length; i++) {
        var currentParagraph = paragraphs[i];

        // Titel-Absatzformat anwenden, wenn es der erste Absatz ist (Titel)
        if (i === 0) {
          currentParagraph.appliedParagraphStyle = findParagraphStyleByPath(
            doc,
            titleParagraphStyleDropdown.selection.text
          );
        } else {
          // Eintrags-Absatzformat für alle anderen Absätze anwenden
          currentParagraph.appliedParagraphStyle = findParagraphStyleByPath(
            doc,
            entryParagraphStyleDropdown.selection.text
          );
        }
      }

      // Label speichern
      selectedTextFrame.label = JSON.stringify({
        type: "ChapterOverviewFrame", // Identifikator für den Textrahmen
        paragraphStyle: selectedStylePath,
        title: customTitle,
        prefixText: prefixText,
        separatorText: separatorText,
        formatType: radioButtonChapterOverview.value
          ? "chapterOverview"
          : "classicTOC",
        titleParagraphStyle: titleParagraphStyleDropdown.selection.text, // Absatzformat für den Titel
        entryParagraphStyle: entryParagraphStyleDropdown.selection.text, // Absatzformat für die Einträge
      });
      $.writeln("Label gespeichert: " + selectedTextFrame.label); // Debug-Ausgabe
    } else {
      // Neuen Textrahmen auf der aktiven Seite erstellen
      var activePage = app.activeWindow.activePage;
      var newTextFrame = activePage.textFrames.add();

      newTextFrame.label = JSON.stringify({
        type: "ChapterOverviewFrame", // Identifikator für den Textrahmen
        paragraphStyle: selectedStylePath,
        title: customTitle,
        prefixText: prefixText,
        separatorText: separatorText,
        formatType: radioButtonChapterOverview.value
          ? "chapterOverview"
          : "classicTOC",
        titleParagraphStyle: titleParagraphStyleDropdown.selection.text, // Absatzformat für den Titel
        entryParagraphStyle: entryParagraphStyleDropdown.selection.text, // Absatzformat für die Einträge
      });
      $.writeln("Label gespeichert: " + newTextFrame.label); // Debug-Ausgabe

      // Textinhalt und Geometrie definieren

      newTextFrame.contents = outputText;

      // Textabsätze formatieren
      var paragraphs = newTextFrame.paragraphs;
      for (var i = 0; i < paragraphs.length; i++) {
        var currentParagraph = paragraphs[i];

        // Titel-Absatzformat anwenden, wenn es der erste Absatz ist (Titel)
        if (i === 0) {
          currentParagraph.appliedParagraphStyle = findParagraphStyleByPath(
            doc,
            titleParagraphStyleDropdown.selection.text
          );
        } else {
          // Eintrags-Absatzformat für alle anderen Absätze anwenden
          currentParagraph.appliedParagraphStyle = findParagraphStyleByPath(
            doc,
            entryParagraphStyleDropdown.selection.text
          );
        }
      }

      // Textrahmen-Positionierung
      var top = "20mm";
      var left = "20mm";
      var bottom = "150mm";
      var right = "150mm";
      newTextFrame.geometricBounds = [top, left, bottom, right];
    }
    alert("Inhaltsverzeichnis wurde erstellt.");
  } else {
    // Text im bestehenden Textrahmen aktualisieren

    existingTextFrame.contents = outputText;

    // Textabsätze formatieren
    var paragraphs = existingTextFrame.paragraphs;
    for (var i = 0; i < paragraphs.length; i++) {
      var currentParagraph = paragraphs[i];

      // Titel-Absatzformat anwenden, wenn es der erste Absatz ist (Titel)
      if (i === 0) {
        currentParagraph.appliedParagraphStyle = findParagraphStyleByPath(
          doc,
          titleParagraphStyleDropdown.selection.text
        );
      } else {
        // Eintrags-Absatzformat für alle anderen Absätze anwenden
        currentParagraph.appliedParagraphStyle = findParagraphStyleByPath(
          doc,
          entryParagraphStyleDropdown.selection.text
        );
      }
    }

    // Label aktualisieren
    existingTextFrame.label = JSON.stringify({
      type: "ChapterOverviewFrame", // Identifikator für den Textrahmen
      paragraphStyle: selectedStylePath,
      title: customTitle,
      prefixText: prefixText,
      separatorText: separatorText,
      formatType: radioButtonChapterOverview.value
        ? "chapterOverview"
        : "classicTOC",
      titleParagraphStyle: titleParagraphStyleDropdown.selection.text, // Absatzformat für den Titel
      entryParagraphStyle: entryParagraphStyleDropdown.selection.text, // Absatzformat für die Einträge
    });
    $.writeln("Label aktualisiert: " + existingTextFrame.label); // Debug-Ausgabe
    alert("Inhaltsverzeichnis wurde aktualisiert.");
  }

  // Funktion, um alle Absatzformate im Dokument mit vollständigem Pfad zu erhalten
  function getAllParagraphStylesWithPath(doc) {
    var styles = [];

    function getStylesFromGroup(group, path) {
      for (var i = 0; i < group.paragraphStyles.length; i++) {
        styles.push(path + group.paragraphStyles[i].name);
      }

      for (var j = 0; j < group.paragraphStyleGroups.length; j++) {
        getStylesFromGroup(
          group.paragraphStyleGroups[j],
          path + group.paragraphStyleGroups[j].name + " > "
        );
      }
    }

    getStylesFromGroup(doc, "");
    return styles;
  }

  // Beispiel, um sicherzustellen, dass der Absatzstil existiert und eine Fehlermeldung angezeigt wird
  function findParagraphStyleByPath(doc, stylePath) {
    var pathParts = stylePath.split(" > ");
    var style = doc.paragraphStyles;

    for (var i = 0; i < pathParts.length; i++) {
      style = style.itemByName(pathParts[i]);

      if (!style.isValid) {
        alert(
          "Absatzformat '" + pathParts[i] + "' konnte nicht gefunden werden."
        );
        return null; // Rückgabe von null, wenn der Stil nicht existiert
      }

      if (i < pathParts.length - 1) {
        style = style.paragraphStyleGroups;
      }
    }

    return style;
  }

  // Überprüfen, ob der Stil existiert
  var selectedStylePath = paragraphStyleDropdown.selection.text;
  var chapterStyle = findParagraphStyleByPath(doc, selectedStylePath);

  if (!chapterStyle) {
    alert("Das ausgewählte Absatzformat existiert nicht.");
    return;
  }
}

// Skript ausführen
createChapterPageRange();
