// createChapterPageRange.jsx

// DESCRIPTION:Erstellen von Inhaltsverzeichnissen und Kapitelübersichten
// AUTHOR: Mario Fritsche
// DATE: 30.12.2024
//Version: 1.0

function createChapterPageRange() {
  var doc = app.activeDocument;

  // JSON
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

  var mainGroup = dialog.add("group");
  mainGroup.orientation = "column";
  mainGroup.alignChildren = "left";

  // Panel für Formatwahl
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

  radioButtonTableOfContents.value = true;

  // Panel "Allgemein"
  var generalPanel = mainGroup.add("panel", undefined, "Allgemein");
  generalPanel.alignChildren = "left";
  generalPanel.preferredSize.width = 350;

  generalPanel.add("statictext", undefined, "Titel");
  var titleInput = generalPanel.add(
    "edittext",
    undefined,
    "Inhaltsverzeichnis"
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
  paragraphStyleDropdown.selection = 0;

  generalPanel.add("statictext", undefined, "Text vor der Seitenzahl:");
  var prefixInput = generalPanel.add("edittext", undefined, "Seite");
  prefixInput.characters = 15;

  generalPanel.add("statictext", undefined, "Text zwischen den Seitenzahlen:");
  var separatorInput = generalPanel.add("edittext", undefined, "-");
  separatorInput.characters = 15;
  separatorInput.enabled = false;

  // Panel für Optionen (Absatzformate Titel Einträge)
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

  // Checkbox
  var updateCheckbox = optionsPanel.add(
    "checkbox",
    undefined,
    "Vorhandenes Inhaltsverzeichnis aktualisieren"
  );

  optionsPanel.add(
    "statictext",
    undefined,
    "Wählen Sie ein Inhaltsverzeichnis:"
  );

  var existingFramesDropdown = optionsPanel.add("dropdownlist", undefined, []);
  existingFramesDropdown.selection = 0;
  existingFramesDropdown.enabled = false;
  existingFramesDropdown.preferredSize.width = 150;

  // Erste Zeile eines Textrahmens
  function getFirstLineOfTextFrame(frame) {
    if (!frame || !frame.contents) {
      return "Unbenannter Textrahmen";
    }
    var lines = frame.contents.split("\r");
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
      if (labelData.formatType === "chapterOverview") {
        radioButtonChapterOverview.value = true;
        separatorInput.enabled = true;
      } else if (labelData.formatType === "classicTOC") {
        radioButtonTableOfContents.value = true;
        separatorInput.enabled = false;
      }
      // Titel, Prefix und Separator vorbefüllen
      titleInput.text = labelData.title || "";
      prefixInput.text = labelData.prefixText || "";
      separatorInput.text = labelData.separatorText || "";

      // Absatzformat für das Inhaltsverzeichnis aus dem Label holen
      var selectedStyle = labelData.paragraphStyle;
      var index = paragraphStyleDropdown.find(selectedStyle);

      if (index !== -1) {
        paragraphStyleDropdown.selection = index;
      } else {
        alert(
          "Das Absatzformat '" +
            selectedStyle +
            "' wurde nicht gefunden. Es wird das erste Format ausgewählt."
        );
        paragraphStyleDropdown.selection = 0;
      }

      // Absatzformate für Titel und Einträge
      var titleParagraphStyle = labelData.titleParagraphStyle;
      var entryParagraphStyle = labelData.entryParagraphStyle;

      // Dropdowns für Titel und Einträge
      var titleStyleIndex =
        titleParagraphStyleDropdown.find(titleParagraphStyle);
      if (titleStyleIndex !== -1) {
        titleParagraphStyleDropdown.selection = titleStyleIndex;
      } else {
        titleParagraphStyleDropdown.selection = 0;
      }

      var entryStyleIndex =
        entryParagraphStyleDropdown.find(entryParagraphStyle);
      if (entryStyleIndex !== -1) {
        entryParagraphStyleDropdown.selection = entryStyleIndex;
      } else {
        entryParagraphStyleDropdown.selection = 0;
      }
    }
  };

  // Index Absatzformat
  function findParagraphStyleIndex(dropdown, styleName) {
    for (var i = 0; i < dropdown.items.length; i++) {
      if (dropdown.items[i].text === styleName) {
        return i;
      }
    }
    return -1;
  }

  // Checkbox-Änderung überwachen, um das Dropdown zu aktivieren
  updateCheckbox.onClick = function () {
    existingFramesDropdown.enabled = updateCheckbox.value;
  };

  // Wenn kein Textrahmen mit dem Label "ChapterOverviewFrame" existiert, Update-Checkbox deaktivieren
  if (existingTextFrames.length === 0) {
    updateCheckbox.enabled = false;
  }

  // Event-Handler für den Radio-Button für Inhaltsverzeichnis
  radioButtonTableOfContents.onClick = function () {
    if (radioButtonTableOfContents.value) {
      separatorInput.enabled = false;
    }
  };

  // Event-Handler für den Radio-Button für Kapitelübersicht
  radioButtonChapterOverview.onClick = function () {
    if (radioButtonChapterOverview.value) {
      separatorInput.enabled = true;
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

  // Kapitelübersicht oder das Inhaltsverzeichnis erstellen
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
    alert(
      (radioButtonChapterOverview.value
        ? "Kapitelübersicht"
        : "Inhaltsverzeichnis") + " wurde erstellt."
    );
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

    alert(
      (radioButtonChapterOverview.value
        ? "Kapitelübersicht"
        : "Inhaltsverzeichnis") + " wurde aktualisiert."
    );
  }

  // Alle Absatzformate im Dokument mit vollständigem Pfad
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

  // Absatzformat muss vorhanden sein
  function findParagraphStyleByPath(doc, stylePath) {
    var pathParts = stylePath.split(" > ");
    var style = doc.paragraphStyles;

    for (var i = 0; i < pathParts.length; i++) {
      style = style.itemByName(pathParts[i]);

      if (!style.isValid) {
        alert(
          "Absatzformat '" + pathParts[i] + "' konnte nicht gefunden werden."
        );
        return null;
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
