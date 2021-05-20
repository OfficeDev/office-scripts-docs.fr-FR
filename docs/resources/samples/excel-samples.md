---
title: Scripts de base pour Office scripts en Excel sur le Web
description: Une collection d’échantillons de code à utiliser avec Office scripts en Excel sur le Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f252934a92126212b9520223826b3b2f5161ed57
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545758"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a>Scripts de base pour Office scripts en Excel sur le Web

Les échantillons suivants sont des scripts simples pour vous d’essayer sur vos propres cahiers de travail. Pour les utiliser en Excel sur le Web :

1. Ouvrez l’onglet **Automatiser**.
2. Éditeur de **code de presse**.
3. Appuyez **sur Nouveau** Script dans le volet de tâche de l’éditeur de code.
4. Remplacez l’ensemble du script par l’échantillon de votre choix.
5. Appuyez **sur Exécuter** dans le volet de tâche de l’éditeur de code.

## <a name="script-basics"></a>Bases du script

Ces échantillons démontrent des éléments constitutifs fondamentaux pour Office scripts. Ajoutez-les à vos scripts pour étendre votre solution et résoudre des problèmes communs.

### <a name="read-and-log-one-cell"></a>Lire et enregistrer une cellule

Cet échantillon lit la valeur de **A1 et** l’imprime sur la console.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a>Lire la cellule active

Ce script enregistre la valeur de la cellule active actuelle. Si plusieurs cellules sont sélectionnées, la cellule la plus à gauche sera enregistrée.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Changer une cellule adjacente

Ce script obtient des cellules adjacentes en utilisant des références relatives. Notez que si la cellule active est sur la ligne supérieure, une partie du script échoue, car elle fait référence à la cellule au-dessus de la cellule actuellement sélectionnée.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a>Changer toutes les cellules adjacentes

Ce script copie le formatage dans la cellule active aux cellules voisines. Notez que ce script ne fonctionne que lorsque la cellule active n’est pas sur un bord de la feuille de travail.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="change-each-individual-cell-in-a-range"></a>Changer chaque cellule dans une plage

Ce script boucle sur la plage actuellement sélectionnée. Il efface le formatage actuel et définit la couleur de remplissage dans chaque cellule à une couleur aléatoire.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();

  // Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();

  // Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;

      // Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

### <a name="get-groups-of-cells-based-on-special-criteria"></a>Obtenez des groupes de cellules en fonction de critères spéciaux

Ce script obtient toutes les cellules vierges dans la plage utilisée de la feuille de travail actuelle. Il met ensuite en évidence toutes ces cellules avec un fond jaune.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a>Collections

Ces échantillons fonctionnent avec des collections d’objets dans le cahier de travail.

### <a name="iterate-over-collections"></a>Itérer sur les collections

Ce script obtient et enregistre les noms de toutes les feuilles de travail dans le cahier de travail. Il définit également les couleurs de leurs onglets à une couleur aléatoire.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();

  // Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());

  // Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;

    // Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

### <a name="query-and-delete-from-a-collection"></a>Requête et suppression d’une collection

Ce script crée une nouvelle feuille de travail. Il vérifie une copie existante de la feuille de travail et la supprime avant de faire une nouvelle feuille.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";

  // Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");

  // Switch to the new worksheet.
  newSheet.activate();
}
```

## <a name="dates"></a>Dates

Les échantillons de cette section montrent comment utiliser l’objet JavaScript [Date.](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date)

L’échantillon suivant obtient la date et l’heure actuelles, puis écrit ces valeurs à deux cellules dans la feuille de travail active.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

L’échantillon suivant lit une date stockée dans Excel et la traduit en objet JavaScript Date. Il utilise le [numéro de série numérique de la date comme](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) entrée pour la date JavaScript.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a>Afficher les données

Ces échantillons démontrent comment travailler avec les données de la feuille de travail et fournissent aux utilisateurs une meilleure vue ou une meilleure organisation.

### <a name="apply-conditional-formatting"></a>Application d’une mise en forme conditionnelle

Cet exemple applique le formatage conditionnel à la plage actuellement utilisée dans la feuille de travail. Le formatage conditionnel est un remplissage vert pour les 10% de valeurs les plus élevées.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a>Créer une table triée

Cet exemple crée une table à partir de la plage utilisée de la feuille de travail actuelle, puis la trie en fonction de la première colonne.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Enregistrez les valeurs « Grand Total » à partir d’un PivotTable

Cet exemple trouve le premier PivotTable dans le cahier de travail et enregistre les valeurs dans les cellules « Grand Total » (comme mis en évidence en vert dans l’image ci-dessous).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="Un PivotTable affichant les ventes de fruits avec la ligne Grand Total mis en évidence vert":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

### <a name="create-a-drop-down-list-using-data-validation"></a>Créer une liste de drop-down à l’aide de la validation des données

Ce script crée une liste de sélection décrocheuse pour une cellule. Il utilise les valeurs existantes de la plage sélectionnée comme choix pour la liste.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Une feuille de travail montrant une gamme de trois cellules contenant des choix de couleurs « rouge, bleu, vert » et à côté de lui, les mêmes choix indiqués dans une liste de drop-down":::

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the values for data validation.
  let selectedRange = workbook.getSelectedRange();
  let rangeValues = selectedRange.getValues();

  // Convert the values into a comma-delimited string.
  let dataValidationListString = "";
  rangeValues.forEach((rangeValueRow) => {
    rangeValueRow.forEach((value) => {
      dataValidationListString += value + ",";
    });
  });

  // Clear the old range.
  selectedRange.clear(ExcelScript.ClearApplyTo.contents);

  // Apply the data validation to the first cell in the selected range.
  let targetCell = selectedRange.getCell(0,0);
  let dataValidation = targetCell.getDataValidation();

  // Set the content of the drop-down list.
  dataValidation.setRule({
      list: {
        inCellDropDown: true,
        source: dataValidationListString
      }
    });
}
```

## <a name="formulas"></a>Formules

Ces échantillons utilisent Excel formules et montrent comment travailler avec eux dans les scripts.

### <a name="single-formula"></a>Formule unique

Ce script définit la formule d’une cellule, puis affiche Excel stocke la formule et la valeur de la cellule séparément.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  // Set A1 to 2.
  let a1 = selectedSheet.getRange("A1");
  a1.setValue(2);

  // Set B1 to the formula =(2*A1), which should equal 4.
  let b1 = selectedSheet.getRange("B1")
  b1.setFormula("=(2*A1)");

  // Log the current results for `getFormula` and `getValue` at B1.
  console.log(`B1 - Formula: ${b1.getFormula()} | Value: ${b1.getValue()}`);
}
```

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Gérer une erreur `#SPILL!` retournée à partir d’une formule

Ce script transpose la gamme « A1:D2 » à « A4:B7 » en utilisant la fonction TRANSPOSE. Si la transpose entraîne une `#SPILL` erreur, elle efface la plage cible et applique à nouveau la formule.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  // Use the data in A1:D2 for the sample.
  let dataAddress = "A1:D2"
  let inputRange = sheet.getRange(dataAddress);

  // Place the transposed data starting at A4.
  let targetStartCell = sheet.getRange("A4");

  // Compute the target range.
  let targetRange = targetStartCell.getResizedRange(inputRange.getColumnCount() - 1, inputRange.getRowCount() - 1);

  // Call the transpose helper function.
  targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);

  // Check if the range update resulted in a spill error.
  let checkValue = targetStartCell.getValue() as string;
  if (checkValue === '#SPILL!') {
    // Clear the target range and call the transpose function again.
    console.log("Target range has data that is preventing update. Clearing target range.");
    targetRange.clear();
    targetStartCell.setFormula(`=TRANSPOSE(${dataAddress})`);
  }

  // Select the transposed range to highlight it.
  targetRange.select();
}
```

## <a name="suggest-new-samples"></a>Suggérer de nouveaux échantillons

Nous accueillons favorablement les suggestions de nouveaux échantillons. S’il existe un scénario commun qui aiderait d’autres développeurs de scripts, veuillez nous le dire dans la section feedback au bas de la page.

## <a name="see-also"></a>Voir aussi

* [Sudhi Ramamurthy « Range basics » sur YouTube](https://youtu.be/4emjkOFdLBA)
* [Office Scripts échantillons et scénarios](samples-overview.md)
* [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](../../tutorials/excel-tutorial.md)
