---
title: Exemples de scripts pour les scripts Office dans Excel sur le web
description: Collection d’exemples de code à utiliser avec les scripts Office dans Excel sur le web.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 4f1f6d4e160c42524df3c69228d182f1cb4838c8
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837273"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="33771-103">Exemples de scripts pour les scripts Office dans Excel sur le web (aperçu)</span><span class="sxs-lookup"><span data-stu-id="33771-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="33771-104">Les exemples suivants sont des scripts simples que vous pouvez essayer sur vos propres workbooks.</span><span class="sxs-lookup"><span data-stu-id="33771-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="33771-105">Pour les utiliser dans Excel sur le web :</span><span class="sxs-lookup"><span data-stu-id="33771-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="33771-106">Ouvrez l’onglet **Automatiser**.</span><span class="sxs-lookup"><span data-stu-id="33771-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="33771-107">Appuyez **sur Éditeur de code.**</span><span class="sxs-lookup"><span data-stu-id="33771-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="33771-108">Appuyez **sur Nouveau script** dans le volet Des tâches de l’Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="33771-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="33771-109">Remplacez l’intégralité du script par l’exemple de votre choix.</span><span class="sxs-lookup"><span data-stu-id="33771-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="33771-110">Appuyez **sur Exécuter** dans le volet Des tâches de l’Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="33771-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="33771-111">Informations de base sur les scripts</span><span class="sxs-lookup"><span data-stu-id="33771-111">Scripting basics</span></span>

<span data-ttu-id="33771-112">Ces exemples montrent les blocs de construction fondamentaux pour les scripts Office.</span><span class="sxs-lookup"><span data-stu-id="33771-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="33771-113">Ajoutez-les à vos scripts pour étendre votre solution et résoudre les problèmes courants.</span><span class="sxs-lookup"><span data-stu-id="33771-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="33771-114">Lire et enregistrer une cellule</span><span class="sxs-lookup"><span data-stu-id="33771-114">Read and log one cell</span></span>

<span data-ttu-id="33771-115">Cet exemple lit la valeur de **A1** et l’imprime sur la console.</span><span class="sxs-lookup"><span data-stu-id="33771-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="33771-116">Lire la cellule active</span><span class="sxs-lookup"><span data-stu-id="33771-116">Read the active cell</span></span>

<span data-ttu-id="33771-117">Ce script enregistre la valeur de la cellule active active.</span><span class="sxs-lookup"><span data-stu-id="33771-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="33771-118">Si plusieurs cellules sont sélectionnées, la cellule située le plus à gauche est enregistrée.</span><span class="sxs-lookup"><span data-stu-id="33771-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="33771-119">Modifier une cellule adjacente</span><span class="sxs-lookup"><span data-stu-id="33771-119">Change an adjacent cell</span></span>

<span data-ttu-id="33771-120">Ce script obtient des cellules adjacentes à l’aide de références relatives.</span><span class="sxs-lookup"><span data-stu-id="33771-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="33771-121">Notez que si la cellule active se trouve sur la ligne supérieure, une partie du script échoue, car elle fait référence à la cellule au-dessus de la cellule actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="33771-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="33771-122">Modifier toutes les cellules adjacentes</span><span class="sxs-lookup"><span data-stu-id="33771-122">Change all adjacent cells</span></span>

<span data-ttu-id="33771-123">Ce script copie la mise en forme de la cellule active vers les cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="33771-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="33771-124">Notez que ce script fonctionne uniquement lorsque la cellule active n’est pas sur un bord de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="33771-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="33771-125">Modifier chaque cellule individuelle d’une plage</span><span class="sxs-lookup"><span data-stu-id="33771-125">Change each individual cell in a range</span></span>

<span data-ttu-id="33771-126">Ce script s’écrit en boucle sur la plage actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="33771-126">This script loops over the currently select range.</span></span> <span data-ttu-id="33771-127">Elle permet d’effacer la mise en forme actuelle et de mettre en couleur aléatoire la couleur de remplissage de chaque cellule.</span><span class="sxs-lookup"><span data-stu-id="33771-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

```typescript
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

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="33771-128">Obtenir des groupes de cellules en fonction de critères spéciaux</span><span class="sxs-lookup"><span data-stu-id="33771-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="33771-129">Ce script obtient toutes les cellules vides de la plage utilisée de la feuille de calcul actuelle.</span><span class="sxs-lookup"><span data-stu-id="33771-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="33771-130">Il met ensuite en évidence toutes ces cellules avec un arrière-plan jaune.</span><span class="sxs-lookup"><span data-stu-id="33771-130">It then highlights all those cells with a yellow background.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);

    // Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## <a name="collections"></a><span data-ttu-id="33771-131">Collections</span><span class="sxs-lookup"><span data-stu-id="33771-131">Collections</span></span>

<span data-ttu-id="33771-132">Ces exemples fonctionnent avec des collections d’objets dans le workbook.</span><span class="sxs-lookup"><span data-stu-id="33771-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterating-over-collections"></a><span data-ttu-id="33771-133">Iterating over collections</span><span class="sxs-lookup"><span data-stu-id="33771-133">Iterating over collections</span></span>

<span data-ttu-id="33771-134">Ce script obtient et enregistre les noms de toutes les feuilles de calcul du manuel.</span><span class="sxs-lookup"><span data-stu-id="33771-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="33771-135">Il définit également les couleurs de leur onglet sur une couleur aléatoire.</span><span class="sxs-lookup"><span data-stu-id="33771-135">It also sets the their tab colors to a random color.</span></span>

```typescript
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

### <a name="querying-and-deleting-from-a-collection"></a><span data-ttu-id="33771-136">Interrogation et suppression d’une collection</span><span class="sxs-lookup"><span data-stu-id="33771-136">Querying and deleting from a collection</span></span>

<span data-ttu-id="33771-137">Ce script crée une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="33771-137">This script creates a new worksheet.</span></span> <span data-ttu-id="33771-138">Il recherche une copie existante de la feuille de calcul et la supprime avant d’en faire une nouvelle.</span><span class="sxs-lookup"><span data-stu-id="33771-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

```typescript
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

## <a name="dates"></a><span data-ttu-id="33771-139">Dates</span><span class="sxs-lookup"><span data-stu-id="33771-139">Dates</span></span>

<span data-ttu-id="33771-140">Les exemples de cette section montrent comment utiliser l’objet [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.</span><span class="sxs-lookup"><span data-stu-id="33771-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="33771-141">L’exemple suivant obtient la date et l’heure actuelles, puis écrit ces valeurs dans deux cellules de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="33771-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="33771-142">L’exemple suivant lit une date stockée dans Excel et la traduit en objet Date JavaScript.</span><span class="sxs-lookup"><span data-stu-id="33771-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="33771-143">Il utilise le [numéro de série numérique](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) de la date comme entrée pour la date JavaScript.</span><span class="sxs-lookup"><span data-stu-id="33771-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="33771-144">Afficher les données</span><span class="sxs-lookup"><span data-stu-id="33771-144">Display data</span></span>

<span data-ttu-id="33771-145">Ces exemples montrent comment travailler avec des données de feuille de calcul et fournir aux utilisateurs une meilleure vue ou une meilleure organisation.</span><span class="sxs-lookup"><span data-stu-id="33771-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="33771-146">Application d’une mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="33771-146">Apply conditional formatting</span></span>

<span data-ttu-id="33771-147">Cet exemple applique une mise en forme conditionnelle à la plage actuellement utilisée dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="33771-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="33771-148">La mise en forme conditionnelle est un remplissage vert pour les 10 % de valeurs les plus importantes.</span><span class="sxs-lookup"><span data-stu-id="33771-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="33771-149">Créer un tableau trié</span><span class="sxs-lookup"><span data-stu-id="33771-149">Create a sorted table</span></span>

<span data-ttu-id="33771-150">Cet exemple crée un tableau à partir de la plage utilisée de la feuille de calcul actuelle, puis le trie en fonction de la première colonne.</span><span class="sxs-lookup"><span data-stu-id="33771-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="33771-151">Enregistrer les valeurs « Total total » à partir d’un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="33771-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="33771-152">Cet exemple recherche le premier tableau croisé dynamique dans le manuel et enregistre les valeurs dans les cellules « Grand Total » (comme indiqué en vert dans l’image ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="33771-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![Tableau croisé dynamique ventes de fruit avec la ligne Total grand mis en surbrillante en vert.](../images/sample-pivottable-grand-total-row.png)

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

### <a name="use-data-validation-to-create-a-drop-down-list"></a><span data-ttu-id="33771-154">Utiliser la validation des données pour créer une liste de listes</span><span class="sxs-lookup"><span data-stu-id="33771-154">Use data validation to create a drop-down list</span></span>

<span data-ttu-id="33771-155">Ce script crée une liste de sélection pour une cellule.</span><span class="sxs-lookup"><span data-stu-id="33771-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="33771-156">Il utilise les valeurs existantes de la plage sélectionnée comme choix pour la liste.</span><span class="sxs-lookup"><span data-stu-id="33771-156">It uses the existing values of the selected range as the choices for the list.</span></span>

![Ensemble de captures d’écran avant et après qui montre trois mots dans une plage, puis ces mêmes mots dans une liste de listes.](../images/sample-data-validation.png)

```typescript
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

## <a name="formulas"></a><span data-ttu-id="33771-158">Formules</span><span class="sxs-lookup"><span data-stu-id="33771-158">Formulas</span></span>

<span data-ttu-id="33771-159">Ces exemples utilisent des formules Excel et montrent comment les utiliser dans des scripts.</span><span class="sxs-lookup"><span data-stu-id="33771-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="33771-160">Formule unique</span><span class="sxs-lookup"><span data-stu-id="33771-160">Single formula</span></span>

<span data-ttu-id="33771-161">Ce script définit la formule d’une cellule, puis montre comment Excel stocke la formule et la valeur de la cellule séparément.</span><span class="sxs-lookup"><span data-stu-id="33771-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

```typescript
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

### <a name="spilling-results-from-a-formula"></a><span data-ttu-id="33771-162">Débordement de résultats d’une formule</span><span class="sxs-lookup"><span data-stu-id="33771-162">Spilling results from a formula</span></span>

<span data-ttu-id="33771-163">Ce script transpose la plage « A1:D2 » en « A4:B7 » à l’aide de la fonction TRANSPOSE.</span><span class="sxs-lookup"><span data-stu-id="33771-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="33771-164">Si la transpose entraîne une erreur #SPILL, elle permet d’effacer la plage cible et d’appliquer à nouveau la formule.</span><span class="sxs-lookup"><span data-stu-id="33771-164">If the transpose results in a #SPILL error, it clears the target range and applies the formula again.</span></span>

```typescript
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

## <a name="scenario-samples"></a><span data-ttu-id="33771-165">Exemples de scénario</span><span class="sxs-lookup"><span data-stu-id="33771-165">Scenario samples</span></span>

<span data-ttu-id="33771-166">Pour obtenir des exemples présentant des solutions réelles et plus volumineuses, consultez les [exemples de scénarios pour Les scripts Office.](scenarios/sample-scenario-overview.md)</span><span class="sxs-lookup"><span data-stu-id="33771-166">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="33771-167">Suggérer de nouveaux exemples</span><span class="sxs-lookup"><span data-stu-id="33771-167">Suggest new samples</span></span>

<span data-ttu-id="33771-168">Nous vous souhaitons la bienvenue pour les nouveaux exemples.</span><span class="sxs-lookup"><span data-stu-id="33771-168">We welcome suggestions for new samples.</span></span> <span data-ttu-id="33771-169">S’il existe un scénario courant qui pourrait aider d’autres développeurs de scripts, n’hésitez pas à nous en faire part dans la section commentaires ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="33771-169">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
