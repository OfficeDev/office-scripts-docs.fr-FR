---
title: Exemples de scripts pour les scripts Office dans Excel sur le Web
description: Collection d’exemples de code à utiliser avec des scripts Office dans Excel sur le Web.
ms.date: 06/18/2020
localization_priority: Normal
ms.openlocfilehash: bfa6679595e6e28cc5d2ae3e3e487fd3e77738aa
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878674"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="9673f-103">Exemples de scripts pour les scripts Office dans Excel sur le Web (aperçu)</span><span class="sxs-lookup"><span data-stu-id="9673f-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="9673f-104">Les exemples suivants sont des scripts simples que vous pouvez essayer dans vos propres classeurs.</span><span class="sxs-lookup"><span data-stu-id="9673f-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="9673f-105">Pour les utiliser dans Excel sur le Web :</span><span class="sxs-lookup"><span data-stu-id="9673f-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="9673f-106">Ouvrez l’onglet **Automatiser**.</span><span class="sxs-lookup"><span data-stu-id="9673f-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="9673f-107">Appuyez sur **éditeur de code**.</span><span class="sxs-lookup"><span data-stu-id="9673f-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="9673f-108">Appuyez sur **nouveau script** dans le volet Office de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="9673f-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="9673f-109">Remplacez l’intégralité du script par l’exemple de votre choix.</span><span class="sxs-lookup"><span data-stu-id="9673f-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="9673f-110">Appuyez sur **exécuter** dans le volet Office de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="9673f-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="9673f-111">Concepts de base des scripts</span><span class="sxs-lookup"><span data-stu-id="9673f-111">Scripting basics</span></span>

<span data-ttu-id="9673f-112">Ces exemples illustrent des blocs de construction fondamentaux pour les scripts Office.</span><span class="sxs-lookup"><span data-stu-id="9673f-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="9673f-113">Ajoutez-les à vos scripts pour étendre votre solution et résoudre les problèmes courants.</span><span class="sxs-lookup"><span data-stu-id="9673f-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="9673f-114">Lecture et journalisation d’une cellule</span><span class="sxs-lookup"><span data-stu-id="9673f-114">Read and log one cell</span></span>

<span data-ttu-id="9673f-115">Cet exemple lit la valeur de **a1** et l’imprime sur la console.</span><span class="sxs-lookup"><span data-stu-id="9673f-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="read-the-active-cell"></a><span data-ttu-id="9673f-116">Lire la cellule active</span><span class="sxs-lookup"><span data-stu-id="9673f-116">Read the active cell</span></span>

<span data-ttu-id="9673f-117">Ce script journalise la valeur de la cellule active active.</span><span class="sxs-lookup"><span data-stu-id="9673f-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="9673f-118">Si plusieurs cellules sont sélectionnées, la cellule située à l’extrême gauche est consignée.</span><span class="sxs-lookup"><span data-stu-id="9673f-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="9673f-119">Modifier une cellule adjacente</span><span class="sxs-lookup"><span data-stu-id="9673f-119">Change an adjacent cell</span></span>

<span data-ttu-id="9673f-120">Ce script obtient des cellules adjacentes à l’aide de références relatives.</span><span class="sxs-lookup"><span data-stu-id="9673f-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="9673f-121">Notez que si la cellule active se trouve sur la ligne supérieure, une partie du script échoue, car elle fait référence à la cellule située au-dessus de la cellule actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="9673f-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

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

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="9673f-122">Modifier toutes les cellules adjacentes</span><span class="sxs-lookup"><span data-stu-id="9673f-122">Change all adjacent cells</span></span>

<span data-ttu-id="9673f-123">Ce script copie la mise en forme de la cellule active vers les cellules voisines.</span><span class="sxs-lookup"><span data-stu-id="9673f-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="9673f-124">Notez que ce script ne fonctionne que lorsque la cellule active ne se trouve pas sur un bord de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="9673f-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="9673f-125">Utiliser des dates</span><span class="sxs-lookup"><span data-stu-id="9673f-125">Work with dates</span></span>

<span data-ttu-id="9673f-126">Les exemples de cette section indiquent comment utiliser l’objet [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9673f-126">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="9673f-127">L’exemple suivant obtient la date et l’heure actuelles, puis écrit ces valeurs dans deux cellules de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="9673f-127">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="9673f-128">L’exemple suivant lit une date stockée dans Excel et la convertit en un objet JavaScript date.</span><span class="sxs-lookup"><span data-stu-id="9673f-128">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="9673f-129">Il utilise le [numéro de série numérique de la date](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) comme entrée pour la date JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9673f-129">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue();
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="9673f-130">Afficher les données</span><span class="sxs-lookup"><span data-stu-id="9673f-130">Display data</span></span>

<span data-ttu-id="9673f-131">Ces exemples montrent comment utiliser les données de feuille de calcul et fournir aux utilisateurs une meilleure vue ou organisation.</span><span class="sxs-lookup"><span data-stu-id="9673f-131">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="9673f-132">Application d’une mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="9673f-132">Apply conditional formatting</span></span>

<span data-ttu-id="9673f-133">Cet exemple applique la mise en forme conditionnelle à la plage utilisée dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="9673f-133">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="9673f-134">La mise en forme conditionnelle est un remplissage vert pour les 10% de valeurs les plus fréquentes.</span><span class="sxs-lookup"><span data-stu-id="9673f-134">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="9673f-135">Créer un tableau trié</span><span class="sxs-lookup"><span data-stu-id="9673f-135">Create a sorted table</span></span>

<span data-ttu-id="9673f-136">Cet exemple montre comment créer un tableau à partir de la plage utilisée dans la feuille de calcul active, puis comment le trier en fonction de la première colonne.</span><span class="sxs-lookup"><span data-stu-id="9673f-136">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="9673f-137">Enregistrer les valeurs « total général » à partir d’un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="9673f-137">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="9673f-138">Cet exemple recherche le premier tableau croisé dynamique dans le classeur et enregistre les valeurs dans les cellules « total général » (comme mise en surbrillance en vert dans l’image ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="9673f-138">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![Tableau croisé dynamique sur les ventes de fruit avec la ligne de total général mise en évidence de vert.](../images/sample-pivottable-grand-total-row.png)

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getRangeBetweenHeaderAndTotal();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## <a name="scenario-samples"></a><span data-ttu-id="9673f-140">Exemples de scénario</span><span class="sxs-lookup"><span data-stu-id="9673f-140">Scenario samples</span></span>

<span data-ttu-id="9673f-141">Pour obtenir des exemples illustrant des solutions plus étendues dans le monde réel, consultez [exemples de scénarios pour les scripts Office](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="9673f-141">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="9673f-142">Suggérer de nouveaux exemples</span><span class="sxs-lookup"><span data-stu-id="9673f-142">Suggest new samples</span></span>

<span data-ttu-id="9673f-143">Nous vous invitons à suggérer de nouveaux exemples.</span><span class="sxs-lookup"><span data-stu-id="9673f-143">We welcome suggestions for new samples.</span></span> <span data-ttu-id="9673f-144">S’il existe un scénario courant qui aide les autres développeurs de script, veuillez nous en indiquer dans la section commentaires ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="9673f-144">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
