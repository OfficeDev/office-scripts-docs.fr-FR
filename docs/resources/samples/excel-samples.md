---
title: Scripts de base pour Office scripts dans Excel sur le Web
description: Collection d’exemples de code à utiliser avec Office scripts dans Excel sur le Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 3bf3bd5acd10bc5999db4746a2ed62af85237e48
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074556"
---
# <a name="basic-scripts-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="74b05-103">Scripts de base pour Office scripts dans Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="74b05-103">Basic scripts for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="74b05-104">Les exemples suivants sont des scripts simples que vous pouvez essayer sur vos propres workbooks.</span><span class="sxs-lookup"><span data-stu-id="74b05-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="74b05-105">Pour les utiliser dans Excel sur le Web :</span><span class="sxs-lookup"><span data-stu-id="74b05-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="74b05-106">Ouvrez l’onglet **Automatiser**.</span><span class="sxs-lookup"><span data-stu-id="74b05-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="74b05-107">Appuyez **sur Éditeur de code.**</span><span class="sxs-lookup"><span data-stu-id="74b05-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="74b05-108">Appuyez **sur Nouveau script** dans le volet Des tâches de l’Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="74b05-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="74b05-109">Remplacez l’intégralité du script par l’exemple de votre choix.</span><span class="sxs-lookup"><span data-stu-id="74b05-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="74b05-110">Appuyez **sur Exécuter** dans le volet Des tâches de l’Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="74b05-110">Press **Run** in the Code Editor's task pane.</span></span>

## <a name="script-basics"></a><span data-ttu-id="74b05-111">Principes de base des scripts</span><span class="sxs-lookup"><span data-stu-id="74b05-111">Script basics</span></span>

<span data-ttu-id="74b05-112">Ces exemples montrent les blocs de construction fondamentaux pour Office scripts.</span><span class="sxs-lookup"><span data-stu-id="74b05-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="74b05-113">Ajoutez-les à vos scripts pour étendre votre solution et résoudre les problèmes courants.</span><span class="sxs-lookup"><span data-stu-id="74b05-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="74b05-114">Lire et enregistrer une cellule</span><span class="sxs-lookup"><span data-stu-id="74b05-114">Read and log one cell</span></span>

<span data-ttu-id="74b05-115">Cet exemple lit la valeur de **A1** et l’imprime sur la console.</span><span class="sxs-lookup"><span data-stu-id="74b05-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="read-the-active-cell"></a><span data-ttu-id="74b05-116">Lire la cellule active</span><span class="sxs-lookup"><span data-stu-id="74b05-116">Read the active cell</span></span>

<span data-ttu-id="74b05-117">Ce script enregistre la valeur de la cellule active active.</span><span class="sxs-lookup"><span data-stu-id="74b05-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="74b05-118">Si plusieurs cellules sont sélectionnées, la cellule située le plus à gauche est enregistrée.</span><span class="sxs-lookup"><span data-stu-id="74b05-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="74b05-119">Modifier une cellule adjacente</span><span class="sxs-lookup"><span data-stu-id="74b05-119">Change an adjacent cell</span></span>

<span data-ttu-id="74b05-120">Ce script obtient des cellules adjacentes à l’aide de références relatives.</span><span class="sxs-lookup"><span data-stu-id="74b05-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="74b05-121">Notez que si la cellule active se trouve sur la ligne supérieure, une partie du script échoue, car elle fait référence à la cellule au-dessus de la cellule actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="74b05-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="74b05-122">Modifier toutes les cellules adjacentes</span><span class="sxs-lookup"><span data-stu-id="74b05-122">Change all adjacent cells</span></span>

<span data-ttu-id="74b05-123">Ce script copie la mise en forme de la cellule active vers les cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="74b05-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="74b05-124">Notez que ce script fonctionne uniquement lorsque la cellule active n’est pas sur un bord de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="74b05-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

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

### <a name="change-each-individual-cell-in-a-range"></a><span data-ttu-id="74b05-125">Modifier chaque cellule individuelle d’une plage</span><span class="sxs-lookup"><span data-stu-id="74b05-125">Change each individual cell in a range</span></span>

<span data-ttu-id="74b05-126">Ce script s’écrit en boucle sur la plage actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="74b05-126">This script loops over the currently select range.</span></span> <span data-ttu-id="74b05-127">Elle permet d’effacer la mise en forme actuelle et de mettre en couleur aléatoire la couleur de remplissage de chaque cellule.</span><span class="sxs-lookup"><span data-stu-id="74b05-127">It clears the current formatting and sets the fill color in each cell to a random color.</span></span>

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

### <a name="get-groups-of-cells-based-on-special-criteria"></a><span data-ttu-id="74b05-128">Obtenir des groupes de cellules en fonction de critères spéciaux</span><span class="sxs-lookup"><span data-stu-id="74b05-128">Get groups of cells based on special criteria</span></span>

<span data-ttu-id="74b05-129">Ce script obtient toutes les cellules vides de la plage utilisée de la feuille de calcul actuelle.</span><span class="sxs-lookup"><span data-stu-id="74b05-129">This script gets all the blank cells in the current worksheet's used range.</span></span> <span data-ttu-id="74b05-130">Il met ensuite en évidence toutes ces cellules avec un arrière-plan jaune.</span><span class="sxs-lookup"><span data-stu-id="74b05-130">It then highlights all those cells with a yellow background.</span></span>

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

## <a name="collections"></a><span data-ttu-id="74b05-131">Collections</span><span class="sxs-lookup"><span data-stu-id="74b05-131">Collections</span></span>

<span data-ttu-id="74b05-132">Ces exemples fonctionnent avec des collections d’objets dans le workbook.</span><span class="sxs-lookup"><span data-stu-id="74b05-132">These samples work with collections of objects in the workbook.</span></span>

### <a name="iterate-over-collections"></a><span data-ttu-id="74b05-133">Itérer sur les collections</span><span class="sxs-lookup"><span data-stu-id="74b05-133">Iterate over collections</span></span>

<span data-ttu-id="74b05-134">Ce script obtient et enregistre les noms de toutes les feuilles de calcul du manuel.</span><span class="sxs-lookup"><span data-stu-id="74b05-134">This script gets and logs the names of all the worksheets in the workbook.</span></span> <span data-ttu-id="74b05-135">Il définit également les couleurs de leurs onglets sur une couleur aléatoire.</span><span class="sxs-lookup"><span data-stu-id="74b05-135">It also sets the their tab colors to a random color.</span></span>

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

### <a name="query-and-delete-from-a-collection"></a><span data-ttu-id="74b05-136">Interroger et supprimer d’une collection</span><span class="sxs-lookup"><span data-stu-id="74b05-136">Query and delete from a collection</span></span>

<span data-ttu-id="74b05-137">Ce script crée une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="74b05-137">This script creates a new worksheet.</span></span> <span data-ttu-id="74b05-138">Il recherche une copie existante de la feuille de calcul et la supprime avant d’en faire une nouvelle.</span><span class="sxs-lookup"><span data-stu-id="74b05-138">It checks for an existing copy of the worksheet and deletes it before making a new sheet.</span></span>

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

## <a name="dates"></a><span data-ttu-id="74b05-139">Dates</span><span class="sxs-lookup"><span data-stu-id="74b05-139">Dates</span></span>

<span data-ttu-id="74b05-140">Les exemples de cette section montrent comment utiliser l’objet [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.</span><span class="sxs-lookup"><span data-stu-id="74b05-140">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="74b05-141">L’exemple suivant obtient la date et l’heure actuelles, puis écrit ces valeurs dans deux cellules de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="74b05-141">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="74b05-142">L’exemple suivant lit une date stockée dans Excel et la traduit en objet Date JavaScript.</span><span class="sxs-lookup"><span data-stu-id="74b05-142">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="74b05-143">Il utilise le [numéro de série numérique](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) de la date comme entrée pour la date JavaScript.</span><span class="sxs-lookup"><span data-stu-id="74b05-143">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

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

## <a name="display-data"></a><span data-ttu-id="74b05-144">Afficher les données</span><span class="sxs-lookup"><span data-stu-id="74b05-144">Display data</span></span>

<span data-ttu-id="74b05-145">Ces exemples montrent comment travailler avec des données de feuille de calcul et fournir aux utilisateurs une meilleure vue ou une meilleure organisation.</span><span class="sxs-lookup"><span data-stu-id="74b05-145">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="74b05-146">Application d’une mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="74b05-146">Apply conditional formatting</span></span>

<span data-ttu-id="74b05-147">Cet exemple applique une mise en forme conditionnelle à la plage actuellement utilisée dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="74b05-147">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="74b05-148">La mise en forme conditionnelle est un remplissage vert pour les 10 % de valeurs les plus importantes.</span><span class="sxs-lookup"><span data-stu-id="74b05-148">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="74b05-149">Créer un tableau trié</span><span class="sxs-lookup"><span data-stu-id="74b05-149">Create a sorted table</span></span>

<span data-ttu-id="74b05-150">Cet exemple crée un tableau à partir de la plage utilisée de la feuille de calcul actuelle, puis le trie en fonction de la première colonne.</span><span class="sxs-lookup"><span data-stu-id="74b05-150">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="74b05-151">Enregistrer les valeurs « Total total » à partir d’un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="74b05-151">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="74b05-152">Cet exemple recherche le premier tableau croisé dynamique dans le workbook et enregistre les valeurs dans les cellules « Grand Total » (comme indiqué en vert dans l’image ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="74b05-152">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

:::image type="content" source="../../images/sample-pivottable-grand-total-row.png" alt-text="Tableau croisé dynamique affichant les ventes de fruit avec la ligne Grand Total mise en surbrillante en vert.":::

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

### <a name="create-a-drop-down-list-using-data-validation"></a><span data-ttu-id="74b05-154">Créer une liste de listes listes à l’aide de la validation des données</span><span class="sxs-lookup"><span data-stu-id="74b05-154">Create a drop-down list using data validation</span></span>

<span data-ttu-id="74b05-155">Ce script crée une liste de sélection pour une cellule.</span><span class="sxs-lookup"><span data-stu-id="74b05-155">This script creates a drop-down selection list for a cell.</span></span> <span data-ttu-id="74b05-156">Il utilise les valeurs existantes de la plage sélectionnée comme choix pour la liste.</span><span class="sxs-lookup"><span data-stu-id="74b05-156">It uses the existing values of the selected range as the choices for the list.</span></span>

:::image type="content" source="../../images/sample-data-validation.png" alt-text="Feuille de calcul montrant une plage de trois cellules contenant des choix de couleur « rouge, bleu, vert » et en de côté, les mêmes choix affichés dans une liste liste.":::

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

## <a name="formulas"></a><span data-ttu-id="74b05-158">Formules</span><span class="sxs-lookup"><span data-stu-id="74b05-158">Formulas</span></span>

<span data-ttu-id="74b05-159">Ces exemples utilisent Excel formules et montrent comment les utiliser dans des scripts.</span><span class="sxs-lookup"><span data-stu-id="74b05-159">These samples use Excel formulas and show how to work with them in scripts.</span></span>

### <a name="single-formula"></a><span data-ttu-id="74b05-160">Formule unique</span><span class="sxs-lookup"><span data-stu-id="74b05-160">Single formula</span></span>

<span data-ttu-id="74b05-161">Ce script définit la formule d’une cellule, puis Excel la formule et la valeur de la cellule séparément.</span><span class="sxs-lookup"><span data-stu-id="74b05-161">This script sets a cell's formula, then displays how Excel stores the cell's formula and value separately.</span></span>

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

### <a name="handle-a-spill-error-returned-from-a-formula"></a><span data-ttu-id="74b05-162">Gérer une `#SPILL!` erreur renvoyée par une formule</span><span class="sxs-lookup"><span data-stu-id="74b05-162">Handle a `#SPILL!` error returned from a formula</span></span>

<span data-ttu-id="74b05-163">Ce script transpose la plage « A1:D2 » en « A4:B7 » à l’aide de la fonction TRANSPOSE.</span><span class="sxs-lookup"><span data-stu-id="74b05-163">This script transposes the range "A1:D2" to "A4:B7" by using the TRANSPOSE function.</span></span> <span data-ttu-id="74b05-164">Si la transpose entraîne une erreur, elle permet d’effacer la plage `#SPILL` cible et d’appliquer à nouveau la formule.</span><span class="sxs-lookup"><span data-stu-id="74b05-164">If the transpose results in a `#SPILL` error, it clears the target range and applies the formula again.</span></span>

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

## <a name="suggest-new-samples"></a><span data-ttu-id="74b05-165">Suggérer de nouveaux exemples</span><span class="sxs-lookup"><span data-stu-id="74b05-165">Suggest new samples</span></span>

<span data-ttu-id="74b05-166">Nous vous proposons des suggestions bienvenues pour les nouveaux exemples.</span><span class="sxs-lookup"><span data-stu-id="74b05-166">We welcome suggestions for new samples.</span></span> <span data-ttu-id="74b05-167">S’il existe un scénario courant qui pourrait aider d’autres développeurs de scripts, n’hésitez pas à nous en faire part dans la section commentaires en bas de la page.</span><span class="sxs-lookup"><span data-stu-id="74b05-167">If there is a common scenario that would help other script developers, please tell us in the feedback section at the bottom of the page.</span></span>

## <a name="see-also"></a><span data-ttu-id="74b05-168">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="74b05-168">See also</span></span>

* [<span data-ttu-id="74b05-169">« Principes de base de la plage » de Sudhi Journal sur YouTube</span><span class="sxs-lookup"><span data-stu-id="74b05-169">Sudhi Ramamurthy's "Range basics" on YouTube</span></span>](https://youtu.be/4emjkOFdLBA)
* [<span data-ttu-id="74b05-170">Office Exemples de scripts et scénarios</span><span class="sxs-lookup"><span data-stu-id="74b05-170">Office Scripts samples and scenarios</span></span>](samples-overview.md)
* [<span data-ttu-id="74b05-171">Enregistrer, modifier et créer des scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="74b05-171">Record, edit, and create Office Scripts in Excel on the web</span></span>](../../tutorials/excel-tutorial.md)
