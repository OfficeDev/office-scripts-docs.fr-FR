---
title: Informations de base sur les plages dans les scripts Office
description: Découvrez les principes de base de l’utilisation de l’objet Range dans Office Scripts.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 73eeba086aace6262c624de9074ffb301f6532bd
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571168"
---
# <a name="range-basics"></a><span data-ttu-id="74190-103">Informations de base sur les plages</span><span class="sxs-lookup"><span data-stu-id="74190-103">Range basics</span></span>

<span data-ttu-id="74190-104">`Range` est l’objet de base dans le modèle objet Excel Des scripts Office.</span><span class="sxs-lookup"><span data-stu-id="74190-104">`Range` is the foundational object within the Office Scripts Excel object model.</span></span> <span data-ttu-id="74190-105">[Les API de](/javascript/api/office-scripts/excelscript/excelscript.range) plage permettent d’accéder à la fois aux données et au format disponibles sur la grille et lier d’autres objets clés dans Excel, tels que des feuilles de calcul, des tableaux, des graphiques, etc.</span><span class="sxs-lookup"><span data-stu-id="74190-105">[Range APIs](/javascript/api/office-scripts/excelscript/excelscript.range) allow access to both data and format available on the grid and link other key objects within Excel such as worksheets, tables, charts, etc.</span></span>

<span data-ttu-id="74190-106">Une plage est identifiée à l’aide de son adresse telle que « A1:B4 » ou à l’aide d’un élément nommé, qui est une clé nommée pour un ensemble de cellules donné.</span><span class="sxs-lookup"><span data-stu-id="74190-106">A range is identified using its address such as "A1:B4" or using a named-item, which is a named key for a given set of cells.</span></span> <span data-ttu-id="74190-107">Dans le modèle objet Excel, une cellule et un groupe de cellules sont _appelés_ plage.</span><span class="sxs-lookup"><span data-stu-id="74190-107">In the Excel object model, both a cell and group of cells are referred as _range_.</span></span> <span data-ttu-id="74190-108">`Range` peut contenir des attributs au niveau de la cellule, tels que des données au sein d’une cellule, ainsi que des attributs au niveau des cellules, tels que le format, les bordures, etc.</span><span class="sxs-lookup"><span data-stu-id="74190-108">`Range` can contain cell-level attributes such as data within a cell and also cell and cells-level attributes such as format, borders, etc.</span></span>

<span data-ttu-id="74190-109">`Range` peut également être obtenue via la sélection de l’utilisateur qui se compose d’au moins une cellule.</span><span class="sxs-lookup"><span data-stu-id="74190-109">`Range` can also be obtained via user's selection that consists of at least one cell.</span></span> <span data-ttu-id="74190-110">Lorsque vous interagissez avec la plage, il est important de garder ces relations de cellule et de plage claires.</span><span class="sxs-lookup"><span data-stu-id="74190-110">As you interact with the range, it's important to keep these cell and range relationships clear.</span></span>

<span data-ttu-id="74190-111">Voici l’ensemble principal de getters, setters et autres méthodes utiles le plus souvent utilisées dans les scripts.</span><span class="sxs-lookup"><span data-stu-id="74190-111">Following are the core set of getters, setters, and other useful methods most often used in scripts.</span></span> <span data-ttu-id="74190-112">Il s’agit d’un excellent point de départ pour votre parcours d’API.</span><span class="sxs-lookup"><span data-stu-id="74190-112">This is a great starting point for your API journey.</span></span> <span data-ttu-id="74190-113">Les sections suivantes groupent les méthodes et aident à créer un modèle mental lorsque vous commencez à déverrouiller les `Range` API de l’objet.</span><span class="sxs-lookup"><span data-stu-id="74190-113">The later sections group the methods and help to build a mental model as you begin to unlock the `Range` object's APIs.</span></span>

## <a name="example-scripts"></a><span data-ttu-id="74190-114">Exemples de scripts</span><span class="sxs-lookup"><span data-stu-id="74190-114">Example scripts</span></span>

* [<span data-ttu-id="74190-115">Lecture et écriture de base</span><span class="sxs-lookup"><span data-stu-id="74190-115">Basic read and write</span></span>](#basic-read-and-write)
* [<span data-ttu-id="74190-116">Ajouter une ligne à la fin de la feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="74190-116">Add row at the end of worksheet</span></span>](#add-row-at-the-end-of-worksheet)
* [<span data-ttu-id="74190-117">Effacer le filtre de colonne</span><span class="sxs-lookup"><span data-stu-id="74190-117">Clear column filter</span></span>](clear-table-filter-for-active-cell.md)
* [<span data-ttu-id="74190-118">Colorier chaque cellule avec une couleur unique</span><span class="sxs-lookup"><span data-stu-id="74190-118">Color each cell with unique color</span></span>](#color-each-cell-with-unique-color)
* [<span data-ttu-id="74190-119">Mettre à jour la plage avec des valeurs à l’aide d’un tableau 2D</span><span class="sxs-lookup"><span data-stu-id="74190-119">Update range with values using 2-dimensional (2D) array</span></span>](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a><span data-ttu-id="74190-120">Lecture et écriture de base</span><span class="sxs-lookup"><span data-stu-id="74190-120">Basic read and write</span></span>

```TypeScript
/**
 * This script demonstrates basic read-write operations on the Range object.
 */
function main(workbook: ExcelScript.Workbook) {
  const cell = workbook.getActiveCell();
  const prevValue = cell.getValue();
  if (prevValue) {
      console.log(`Active cell's value is: ${prevValue}`);
  } else {
      console.log("Setting active cell's value..");
      cell.setValue("Sample");
  }

  // Get cell next to the right column and set its value and fill color.
  const nextCell = cell.getOffsetRange(0,1);
  nextCell.setValue("Next cell");
  console.log(`Next cell's address is: ${nextCell.getAddress()}`);
  console.log("Setting fill color and font color of next cell...");
  nextCell.getFormat().getFill().setColor("Magenta");
  nextCell.getFormat().getFill().setColor("Cyan");

  // Get the target range address to update with 2-dimensional value.
  const dataRange = nextCell.getOffsetRange(1, 0).getResizedRange(2, 1);
  const DATA = [
    [10, 7],
    [8, 15],
    [12, 1]
  ];
  console.log(`Updating range ${dataRange.getAddress()} with values: ${DATA}`);
  dataRange.setValues(DATA);

  // Formula range.
  const formulaRange = dataRange.getOffsetRange(3, 0).getRow(0);
  console.log(`Updating formula for range: ${formulaRange.getAddress()}`)
  // Since relative formula is being set, we can set the formula of the entire range to the same value.
  formulaRange.setFormulaR1C1("=SUM(R[-3]C:R[-1]C)");
  console.log(`Updating number format for range: ${formulaRange.getAddress()}`)
  // Since the number format is common to the entire range, we can set it to a common format.
  formulaRange.setNumberFormat("0.00");
  return;
}
```

### <a name="add-row-at-the-end-of-worksheet"></a><span data-ttu-id="74190-121">Ajouter une ligne à la fin de la feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="74190-121">Add row at the end of worksheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('Sheet5');
    const data = ['2016', 'Bikes', 'Seats', '1500', .05];
    addRow(sheet, data);
    return;
}

function addRow(sheet: ExcelScript.Worksheet, data: (string | number | boolean)[]): void {

    const usedRange = sheet.getUsedRange();
    let startCell: ExcelScript.Range;
    // If the sheet is empty, then use A1 as starting cell for the update.
    if (usedRange) {
      startCell = usedRange.getLastRow().getCell(0, 0).getOffsetRange(1, 0);
    } else {
      startCell = sheet.getRange('A1');
    }
    console.log(startCell.getAddress());
    const targetRange = startCell.getResizedRange(0, data.length - 1);
    targetRange.setValues([data]);
    return;
}
```

### <a name="color-each-cell-with-unique-color"></a><span data-ttu-id="74190-122">Colorier chaque cellule avec une couleur unique</span><span class="sxs-lookup"><span data-stu-id="74190-122">Color each cell with unique color</span></span>

```TypeScript
/**
 * This sample demonstrates how to iterate over a selected range and set cell property.
   It colors each cell within the selected range with a random color.
 */
function main(workbook: ExcelScript.Workbook) {

    const syncStart = new Date().getTime();
    // Get selected range
    const range = workbook.getSelectedRange();
    const rows = range.getRowCount();
    const cols = range.getColumnCount();
    console.log("Start");

    // Color each cell with random color.
    for (let row = 0; row < rows; row++) {
        for (let col = 0; col < cols; col++) {
            range
                .getCell(row, col)
                .getFormat()
                .getFill()
                .setColor(`#${Math.random().toString(16).substr(-6)}`);
        }
    }

    console.log("End");
    const syncEnd = new Date().getTime();
    console.log("Completed, took: " + (syncEnd - syncStart) / 1000 + " Sec");
}
```

### <a name="update-range-with-values-using-2d-array"></a><span data-ttu-id="74190-123">Mettre à jour la plage avec des valeurs à l’aide d’un tableau 2D</span><span class="sxs-lookup"><span data-stu-id="74190-123">Update range with values using 2D array</span></span>

<span data-ttu-id="74190-124">Calcule dynamiquement la dimension de plage à mettre à jour en fonction des valeurs de tableau 2D.</span><span class="sxs-lookup"><span data-stu-id="74190-124">Dynamically calculates the range dimension to update based on 2D array values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const currentCell = workbook.getActiveCell();
  let inputRange = computeTargetRange(currentCell, DATA);
  // Set range values.
  console.log(inputRange.getAddress());
  inputRange.setValues(DATA);
  // Call a helper function to place border around the range.
  borderAround(inputRange);
}

/**
 * A helper function that computes the target range given the target range's starting cell and selected range. 
 */
function computeTargetRange(targetCell: ExcelScript.Range, data: string[][]): ExcelScript.Range {
  const targetRange = targetCell.getResizedRange(data.length - 1, data[0].length - 1);
  return targetRange;
}

/**
 * A helper function that places a border around the range.
 */
function borderAround(range: ExcelScript.Range): void {
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.dash);
  range.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.dash);
  return;
}

// Values used for range setup.
const DATA = [
  ['Item', 'Bread', 'Donuts', 'Cookies', 'Cakes', 'Pies'],
  ['Amount', '2', '1.5', '4', '12', '26']
]
```

## <a name="training-videos-range-basics"></a><span data-ttu-id="74190-125">Vidéos de formation : informations de base sur les plages</span><span class="sxs-lookup"><span data-stu-id="74190-125">Training videos: Range basics</span></span>

<span data-ttu-id="74190-126">_Informations de base sur les plages_</span><span class="sxs-lookup"><span data-stu-id="74190-126">_Range basics_</span></span>

<span data-ttu-id="74190-127">[![Regarder une vidéo pas à pas sur les principes de base de la plage](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Vidéo pas à pas sur les principes de base de la plage")</span><span class="sxs-lookup"><span data-stu-id="74190-127">[![Watch step-by-step video on Range basics](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Step-by-step video on Range basics")</span></span>

<span data-ttu-id="74190-128">_Ajouter une ligne à la fin de la feuille de calcul_</span><span class="sxs-lookup"><span data-stu-id="74190-128">_Add row at the end of worksheet_</span></span>

<span data-ttu-id="74190-129">[![Regardez une vidéo détaillée sur l’ajout d’une ligne à la fin d’une feuille de calcul](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Vidéo pas à pas sur l’ajout d’une ligne à la fin d’une feuille de calcul")</span><span class="sxs-lookup"><span data-stu-id="74190-129">[![Watch step-by-step video on how to add a row at the end of a worksheet](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Step-by-step video on how to add a row at the end of a worksheet")</span></span>

## <a name="methods-that-return-some-range-metadata"></a><span data-ttu-id="74190-130">Méthodes qui retournent des métadonnées de plage</span><span class="sxs-lookup"><span data-stu-id="74190-130">Methods that return some range metadata</span></span>

* <span data-ttu-id="74190-131">getAddress(), getAddressLocal()</span><span class="sxs-lookup"><span data-stu-id="74190-131">getAddress(), getAddressLocal()</span></span>
* <span data-ttu-id="74190-132">getCellCount()</span><span class="sxs-lookup"><span data-stu-id="74190-132">getCellCount()</span></span>
* <span data-ttu-id="74190-133">getRowCount(), getColumnCount()</span><span class="sxs-lookup"><span data-stu-id="74190-133">getRowCount(), getColumnCount()</span></span>

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a><span data-ttu-id="74190-134">Méthodes qui retournent des données/constantes associées à une plage donnée</span><span class="sxs-lookup"><span data-stu-id="74190-134">Methods that return data/constants associated with a given range</span></span>

### <a name="returned-as-single-cell-value"></a><span data-ttu-id="74190-135">Renvoyé en tant que valeur de cellule unique</span><span class="sxs-lookup"><span data-stu-id="74190-135">Returned as single cell value</span></span>

* <span data-ttu-id="74190-136">getFormula(), getFormulaLocal()</span><span class="sxs-lookup"><span data-stu-id="74190-136">getFormula(), getFormulaLocal()</span></span>
* <span data-ttu-id="74190-137">getFormulaR1C1()</span><span class="sxs-lookup"><span data-stu-id="74190-137">getFormulaR1C1()</span></span>
* <span data-ttu-id="74190-138">getNumberFormat(), getNumberFormatLocal()</span><span class="sxs-lookup"><span data-stu-id="74190-138">getNumberFormat(), getNumberFormatLocal()</span></span>
* <span data-ttu-id="74190-139">getText()</span><span class="sxs-lookup"><span data-stu-id="74190-139">getText()</span></span>
* <span data-ttu-id="74190-140">getValue()</span><span class="sxs-lookup"><span data-stu-id="74190-140">getValue()</span></span>
* <span data-ttu-id="74190-141">getValueType()</span><span class="sxs-lookup"><span data-stu-id="74190-141">getValueType()</span></span>

### <a name="returned-as-2d-arrays-whole-range"></a><span data-ttu-id="74190-142">Renvoyé en tant que tableaux 2D (plage entière)</span><span class="sxs-lookup"><span data-stu-id="74190-142">Returned as 2D arrays (whole range)</span></span>

* <span data-ttu-id="74190-143">getFormulas(), getFormulasLocal()</span><span class="sxs-lookup"><span data-stu-id="74190-143">getFormulas(), getFormulasLocal()</span></span>
* <span data-ttu-id="74190-144">getFormulasR1C1()</span><span class="sxs-lookup"><span data-stu-id="74190-144">getFormulasR1C1()</span></span>
* <span data-ttu-id="74190-145">getNumberFormatCategories()</span><span class="sxs-lookup"><span data-stu-id="74190-145">getNumberFormatCategories()</span></span>
* <span data-ttu-id="74190-146">getNumberFormats(), getNumberFormatsLocal()</span><span class="sxs-lookup"><span data-stu-id="74190-146">getNumberFormats(), getNumberFormatsLocal()</span></span>
* <span data-ttu-id="74190-147">getTexts()</span><span class="sxs-lookup"><span data-stu-id="74190-147">getTexts()</span></span>
* <span data-ttu-id="74190-148">getValues()</span><span class="sxs-lookup"><span data-stu-id="74190-148">getValues()</span></span>
* <span data-ttu-id="74190-149">getValueTypes()</span><span class="sxs-lookup"><span data-stu-id="74190-149">getValueTypes()</span></span>
* <span data-ttu-id="74190-150">getHidden()</span><span class="sxs-lookup"><span data-stu-id="74190-150">getHidden()</span></span>
* <span data-ttu-id="74190-151">getIsEntireRow()</span><span class="sxs-lookup"><span data-stu-id="74190-151">getIsEntireRow()</span></span>
* <span data-ttu-id="74190-152">getIsEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="74190-152">getIsEntireColumn()</span></span>

## <a name="methods-that-return-other-range-object"></a><span data-ttu-id="74190-153">Méthodes qui retournent un autre objet de plage</span><span class="sxs-lookup"><span data-stu-id="74190-153">Methods that return other range object</span></span>

* <span data-ttu-id="74190-154">getSurroundingRegion() : similaire à CurrentRegion dans VBA</span><span class="sxs-lookup"><span data-stu-id="74190-154">getSurroundingRegion() -- similar to CurrentRegion in VBA</span></span>
* <span data-ttu-id="74190-155">getCell(ligne, colonne)</span><span class="sxs-lookup"><span data-stu-id="74190-155">getCell(row, column)</span></span>
* <span data-ttu-id="74190-156">getColumn(column)</span><span class="sxs-lookup"><span data-stu-id="74190-156">getColumn(column)</span></span>
* <span data-ttu-id="74190-157">getColumnHidden()</span><span class="sxs-lookup"><span data-stu-id="74190-157">getColumnHidden()</span></span>
* <span data-ttu-id="74190-158">getColumnsAfter(count)</span><span class="sxs-lookup"><span data-stu-id="74190-158">getColumnsAfter(count)</span></span>
* <span data-ttu-id="74190-159">getColumnsBefore(count)</span><span class="sxs-lookup"><span data-stu-id="74190-159">getColumnsBefore(count)</span></span>
* <span data-ttu-id="74190-160">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="74190-160">getEntireColumn()</span></span>
* <span data-ttu-id="74190-161">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="74190-161">getEntireRow()</span></span>
* <span data-ttu-id="74190-162">getLastCell()</span><span class="sxs-lookup"><span data-stu-id="74190-162">getLastCell()</span></span>
* <span data-ttu-id="74190-163">getLastColumn()</span><span class="sxs-lookup"><span data-stu-id="74190-163">getLastColumn()</span></span>
* <span data-ttu-id="74190-164">getLastRow()</span><span class="sxs-lookup"><span data-stu-id="74190-164">getLastRow()</span></span>
* <span data-ttu-id="74190-165">getRow(row)</span><span class="sxs-lookup"><span data-stu-id="74190-165">getRow(row)</span></span>
* <span data-ttu-id="74190-166">getRowHidden()</span><span class="sxs-lookup"><span data-stu-id="74190-166">getRowHidden()</span></span>
* <span data-ttu-id="74190-167">getRowsAbove(count)</span><span class="sxs-lookup"><span data-stu-id="74190-167">getRowsAbove(count)</span></span>
* <span data-ttu-id="74190-168">getRowsBelow(count)</span><span class="sxs-lookup"><span data-stu-id="74190-168">getRowsBelow(count)</span></span>

<span data-ttu-id="74190-169">**Important/Intéressant**</span><span class="sxs-lookup"><span data-stu-id="74190-169">**Important/Interesting**</span></span>

* <span data-ttu-id="74190-170">_workbook_.getSelectedRange()</span><span class="sxs-lookup"><span data-stu-id="74190-170">_workbook_.getSelectedRange()</span></span>
* <span data-ttu-id="74190-171">_workbook_.getActiveCell()</span><span class="sxs-lookup"><span data-stu-id="74190-171">_workbook_.getActiveCell()</span></span>
* <span data-ttu-id="74190-172">getUsedRange(valuesOnly)</span><span class="sxs-lookup"><span data-stu-id="74190-172">getUsedRange(valuesOnly)</span></span>
* <span data-ttu-id="74190-173">getAbsoluteResizedRange(numRows, numColumns)</span><span class="sxs-lookup"><span data-stu-id="74190-173">getAbsoluteResizedRange(numRows, numColumns)</span></span>
* <span data-ttu-id="74190-174">getOffsetRange(rowOffset, columnOffset)</span><span class="sxs-lookup"><span data-stu-id="74190-174">getOffsetRange(rowOffset, columnOffset)</span></span>
* <span data-ttu-id="74190-175">getResizedRange(deltaRows, deltaColumns)</span><span class="sxs-lookup"><span data-stu-id="74190-175">getResizedRange(deltaRows, deltaColumns)</span></span>

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a><span data-ttu-id="74190-176">Méthodes qui retournent un objet de plage par rapport à un autre objet de plage</span><span class="sxs-lookup"><span data-stu-id="74190-176">Methods that return a range object in relation to another range object</span></span>

* <span data-ttu-id="74190-177">getBoundingRect(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="74190-177">getBoundingRect(anotherRange)</span></span>
* <span data-ttu-id="74190-178">getIntersection(anotherRange)</span><span class="sxs-lookup"><span data-stu-id="74190-178">getIntersection(anotherRange)</span></span>

## <a name="methods-that-return-other-objects-non-range-objects"></a><span data-ttu-id="74190-179">Méthodes qui retournent d’autres objets (objets autres que des plages)</span><span class="sxs-lookup"><span data-stu-id="74190-179">Methods that return other objects (non-range objects)</span></span>

* <span data-ttu-id="74190-180">getDirectPrecedents()</span><span class="sxs-lookup"><span data-stu-id="74190-180">getDirectPrecedents()</span></span>
* <span data-ttu-id="74190-181">getWorksheet()</span><span class="sxs-lookup"><span data-stu-id="74190-181">getWorksheet()</span></span>
* <span data-ttu-id="74190-182">getTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="74190-182">getTables(fullyContained)</span></span>
* <span data-ttu-id="74190-183">getPivotTables(fullyContained)</span><span class="sxs-lookup"><span data-stu-id="74190-183">getPivotTables(fullyContained)</span></span>
* <span data-ttu-id="74190-184">getDataValidation()</span><span class="sxs-lookup"><span data-stu-id="74190-184">getDataValidation()</span></span>
* <span data-ttu-id="74190-185">getPredefinedCellStyle()</span><span class="sxs-lookup"><span data-stu-id="74190-185">getPredefinedCellStyle()</span></span>

## <a name="set-methods"></a><span data-ttu-id="74190-186">Définir des méthodes</span><span class="sxs-lookup"><span data-stu-id="74190-186">Set methods</span></span>

### <a name="singular-cell-set-methods"></a><span data-ttu-id="74190-187">Méthodes de jeu de cellules au singular</span><span class="sxs-lookup"><span data-stu-id="74190-187">Singular cell set methods</span></span>

* <span data-ttu-id="74190-188">setFormula(formula)</span><span class="sxs-lookup"><span data-stu-id="74190-188">setFormula(formula)</span></span>
* <span data-ttu-id="74190-189">setFormulaLocal(formulaLocal)</span><span class="sxs-lookup"><span data-stu-id="74190-189">setFormulaLocal(formulaLocal)</span></span>
* <span data-ttu-id="74190-190">setFormulaR1C1(formulaR1C1)</span><span class="sxs-lookup"><span data-stu-id="74190-190">setFormulaR1C1(formulaR1C1)</span></span>
* <span data-ttu-id="74190-191">setNumberFormatLocal(numberFormatLocal)</span><span class="sxs-lookup"><span data-stu-id="74190-191">setNumberFormatLocal(numberFormatLocal)</span></span>
* <span data-ttu-id="74190-192">setValue(value)</span><span class="sxs-lookup"><span data-stu-id="74190-192">setValue(value)</span></span>

### <a name="2d--entire-range-set-methods"></a><span data-ttu-id="74190-193">2D / méthodes d’ensemble de plages entières</span><span class="sxs-lookup"><span data-stu-id="74190-193">2D / entire range set methods</span></span>

* <span data-ttu-id="74190-194">setFormulas(formulas)</span><span class="sxs-lookup"><span data-stu-id="74190-194">setFormulas(formulas)</span></span>
* <span data-ttu-id="74190-195">setFormulasLocal(formulasLocal)</span><span class="sxs-lookup"><span data-stu-id="74190-195">setFormulasLocal(formulasLocal)</span></span>
* <span data-ttu-id="74190-196">setFormulasR1C1(formulasR1C1)</span><span class="sxs-lookup"><span data-stu-id="74190-196">setFormulasR1C1(formulasR1C1)</span></span>
* <span data-ttu-id="74190-197">setNumberFormat(numberFormat)</span><span class="sxs-lookup"><span data-stu-id="74190-197">setNumberFormat(numberFormat)</span></span>
* <span data-ttu-id="74190-198">setNumberFormats(numberFormats)</span><span class="sxs-lookup"><span data-stu-id="74190-198">setNumberFormats(numberFormats)</span></span>
* <span data-ttu-id="74190-199">setNumberFormatsLocal(numberFormatsLocal)</span><span class="sxs-lookup"><span data-stu-id="74190-199">setNumberFormatsLocal(numberFormatsLocal)</span></span>
* <span data-ttu-id="74190-200">setValues(values)</span><span class="sxs-lookup"><span data-stu-id="74190-200">setValues(values)</span></span>

## <a name="other-methods"></a><span data-ttu-id="74190-201">Autres méthodes</span><span class="sxs-lookup"><span data-stu-id="74190-201">Other methods</span></span>

* <span data-ttu-id="74190-202">merge(across)</span><span class="sxs-lookup"><span data-stu-id="74190-202">merge(across)</span></span>
* <span data-ttu-id="74190-203">unmerge()</span><span class="sxs-lookup"><span data-stu-id="74190-203">unmerge()</span></span>

## <a name="coming-soon"></a><span data-ttu-id="74190-204">Bientôt disponible</span><span class="sxs-lookup"><span data-stu-id="74190-204">Coming soon</span></span>

* <span data-ttu-id="74190-205">API de bordure de plage</span><span class="sxs-lookup"><span data-stu-id="74190-205">Range edge APIs</span></span>
