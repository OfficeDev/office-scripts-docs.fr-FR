---
title: Scripts de base pour les scripts Office dans Excel
description: Collection d’exemples de code à utiliser avec les scripts Office dans Excel.
ms.date: 06/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 3d17e2cf2314ccd6c07d81e53337fcd63a474fd8
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281902"
---
# <a name="basic-scripts-for-office-scripts-in-excel"></a>Scripts de base pour les scripts Office dans Excel

Les exemples suivants sont des scripts simples que vous pouvez essayer sur vos propres classeurs. Pour les utiliser dans Excel :

1. Ouvrez un classeur dans Excel sur le Web.
1. Ouvrez l’onglet **Automatiser**.
1. Sélectionnez **Nouveau script**.
1. Remplacez l’intégralité du script par l’exemple de votre choix.
1. Sélectionnez **Exécuter** dans le volet Office de l’éditeur de code.

## <a name="script-basics"></a>Concepts de base du script

Ces exemples illustrent les blocs de construction fondamentaux des scripts Office. Développez ces scripts pour étendre votre solution et résoudre les problèmes courants.

### <a name="read-and-log-one-cell"></a>Lire et journaliser une cellule

Cet exemple lit la valeur de **A1** et l’imprime dans la console.

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

Ce script enregistre la valeur de la cellule active actuelle. Si plusieurs cellules sont sélectionnées, la cellule la plus à gauche est journalisée.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a>Modifier une cellule adjacente

Ce script obtient les cellules adjacentes à l’aide de références relatives. Notez que si la cellule active se trouve sur la ligne supérieure, une partie du script échoue, car elle fait référence à la cellule située au-dessus de la cellule actuellement sélectionnée.

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

### <a name="change-all-adjacent-cells"></a>Modifier toutes les cellules adjacentes

Ce script copie la mise en forme de la cellule active dans les cellules voisines. Notez que ce script fonctionne uniquement lorsque la cellule active n’est pas sur un bord de la feuille de calcul.

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

### <a name="change-each-individual-cell-in-a-range"></a>Modifier chaque cellule individuelle d’une plage

Ce script effectue une boucle sur la plage actuellement sélectionnée. Il efface la mise en forme actuelle et définit la couleur de remplissage de chaque cellule sur une couleur aléatoire.

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a>Obtenir des groupes de cellules en fonction de critères spéciaux

Ce script obtient toutes les cellules vides de la plage utilisée de la feuille de calcul actuelle. Il met ensuite en surbrillance toutes ces cellules avec un arrière-plan jaune.

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

### <a name="unhide-all-rows-and-columns"></a>Afficher toutes les lignes et colonnes

Ce script obtient la plage utilisée de la feuille de calcul, vérifie s’il existe des lignes et des colonnes masquées et les annule. 

```Typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the currently selected sheet.
    const selectedSheet = workbook.getActiveWorksheet();

    // Get the entire data range.
    const range = selectedSheet.getUsedRange();

    // If the used range is empty, end the script.
    if (!range) {
      console.log(`No data on this sheet.`)
      return;
    }

    // If no columns are hidden, log message, else, unhide columns
    if (range.getColumnHidden() == false) {
      console.log(`No columns hidden`);
    } else {
      range.setColumnHidden(false);
    }

    // If no rows are hidden, log message, else, unhide rows.
    if (range.getRowHidden() == false) {
      console.log(`No rows hidden`);
    } else {
      range.setRowHidden(false);
    }
}
```

### <a name="freeze-currently-selected-cells"></a>Figer les cellules actuellement sélectionnées

Ce script vérifie les cellules actuellement sélectionnées et fige cette sélection, afin que ces cellules soient toujours visibles.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the currently selected sheet.
    const selectedSheet = workbook.getActiveWorksheet();

    // Get the current selected range.
    const selectedRange = workbook.getSelectedRange();

    // If no cells are selected, end the script. 
    if (!selectedRange) {
      console.log(`No cells in the worksheet are selected.`);
      return;
    }

    // Log the address of the selected range
    console.log(`Selected range for the worksheet: ${selectedRange.getAddress()}`);

    // Freeze the selected range.
    selectedSheet.getFreezePanes().freezeAt(selectedRange);
}
```

## <a name="collections"></a>Collections

Ces exemples fonctionnent avec des collections d’objets dans le classeur.

### <a name="iterate-over-collections"></a>Itérer sur des collections

Ce script obtient et journalise les noms de toutes les feuilles de calcul du classeur. Il définit également les couleurs de leur onglet sur une couleur aléatoire.

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

### <a name="query-and-delete-from-a-collection"></a>Interroger et supprimer d’une collection

Ce script crée une feuille de calcul. Il recherche une copie existante de la feuille de calcul et la supprime avant de créer une nouvelle feuille.

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

Les exemples de cette section montrent comment utiliser [l’objet Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.

L’exemple suivant obtient la date et l’heure actuelles, puis écrit ces valeurs dans deux cellules de la feuille de calcul active.

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

L’exemple suivant lit une date stockée dans Excel et la traduit en objet Date JavaScript. Il utilise le numéro de série numérique de la date comme entrée pour la date JavaScript. Ce numéro de série est décrit dans l’article de [la fonction NOW().](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)

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

Ces exemples montrent comment utiliser des données de feuille de calcul et fournir aux utilisateurs une meilleure vue ou une meilleure organisation.

### <a name="apply-conditional-formatting"></a>Application d’une mise en forme conditionnelle

Cet exemple applique la mise en forme conditionnelle à la plage actuellement utilisée dans la feuille de calcul. La mise en forme conditionnelle est un remplissage vert pour les 10 % de valeurs les plus élevées.

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

Cet exemple crée une table à partir de la plage utilisée de la feuille de calcul actuelle, puis le trie en fonction de la première colonne.

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

### <a name="filter-a-table"></a>Filtrer une table

Cet exemple filtre une table existante à l’aide des valeurs de l’une des colonnes.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table in the workbook named "StationTable".
  const table = workbook.getTable("StationTable");

  // Get the "Station" table column for the filter.
  const stationColumn = table.getColumnByName("Station");

  // Apply a filter to the table that will only show rows 
  // with a value of "Station-1" in the "Station" column.
  stationColumn.getFilter().applyValuesFilter(["Station-1"]);
}
```

> [!TIP]
> Copiez les informations filtrées dans le classeur à l’aide `Range.copyFrom`de . Ajoutez la ligne suivante à la fin du script pour créer une feuille de calcul avec les données filtrées.
>
> ```typescript
>   workbook.addWorksheet().getRange("A1").copyFrom(table.getRange());
> ```

### <a name="log-the-grand-total-values-from-a-pivottable"></a>Journaliser les valeurs « Total général » à partir d’un tableau croisé dynamique

Cet exemple recherche le premier tableau croisé dynamique dans le classeur et consigne les valeurs dans les cellules « Total général » (comme indiqué en vert dans l’image ci-dessous).

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="Tableau croisé dynamique affichant les ventes de fruits avec la ligne Grand Total mise en surbrillance en vert.":::

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

### <a name="create-a-drop-down-list-using-data-validation"></a>Créer une liste déroulante à l’aide de la validation des données

Ce script crée une liste déroulante de sélection pour une cellule. Il utilise les valeurs existantes de la plage sélectionnée comme choix pour la liste.

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Feuille de calcul montrant une plage de trois cellules contenant des choix de couleurs « rouge, bleu, vert » et en regard de celle-ci, les mêmes choix que ceux indiqués dans une liste déroulante.":::

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

Ces exemples utilisent des formules Excel et montrent comment les utiliser dans des scripts.

### <a name="single-formula"></a>Formule unique

Ce script définit la formule d’une cellule, puis affiche la façon dont Excel stocke la formule et la valeur de la cellule séparément.

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a>Gérer une `#SPILL!` erreur retournée à partir d’une formule

Ce script transpose la plage « A1:D2 » en « A4:B7 » à l’aide de la fonction TRANSPOSE. Si la transpose génère une `#SPILL` erreur, elle efface la plage cible et applique à nouveau la formule.

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

### <a name="replace-all-formulas-with-their-result-values"></a>Remplacer toutes les formules par leurs valeurs de résultat

Ce script remplace chaque cellule de la feuille de calcul actuelle qui contient une formule par le résultat de cette formule. Cela signifie qu’il n’y aura pas de formules après l’exécution du script, uniquement des valeurs.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the ranges with formulas.
    let sheet = workbook.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaCells = usedRange.getSpecialCells(ExcelScript.SpecialCellType.formulas);

    // In each formula range: get the current value, clear the contents, and set the value as the old one.
    // This removes the formula but keeps the result.
    formulaCells.getAreas().forEach((range) => {
      let currentValues = range.getValues();
      range.clear(ExcelScript.ClearApplyTo.contents);
      range.setValues(currentValues);
    });
}
```

## <a name="suggest-new-samples"></a>Suggérer de nouveaux exemples

Nous vous invitons à vous proposer de nouveaux exemples. S’il existe un scénario courant qui aiderait d’autres développeurs de scripts, veuillez nous le dire dans la section commentaires en bas de la page.

## <a name="see-also"></a>Voir aussi

* [« Range basics » de Sudhi Ramamurthy sur YouTube](https://youtu.be/4emjkOFdLBA)
* [Exemples et scénarios de scripts Office](samples-overview.md)
* [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](../../tutorials/excel-tutorial.md)
