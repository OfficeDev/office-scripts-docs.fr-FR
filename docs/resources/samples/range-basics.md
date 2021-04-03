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
# <a name="range-basics"></a>Informations de base sur les plages

`Range` est l’objet de base dans le modèle objet Excel Des scripts Office. [Les API de](/javascript/api/office-scripts/excelscript/excelscript.range) plage permettent d’accéder à la fois aux données et au format disponibles sur la grille et lier d’autres objets clés dans Excel, tels que des feuilles de calcul, des tableaux, des graphiques, etc.

Une plage est identifiée à l’aide de son adresse telle que « A1:B4 » ou à l’aide d’un élément nommé, qui est une clé nommée pour un ensemble de cellules donné. Dans le modèle objet Excel, une cellule et un groupe de cellules sont _appelés_ plage. `Range` peut contenir des attributs au niveau de la cellule, tels que des données au sein d’une cellule, ainsi que des attributs au niveau des cellules, tels que le format, les bordures, etc.

`Range` peut également être obtenue via la sélection de l’utilisateur qui se compose d’au moins une cellule. Lorsque vous interagissez avec la plage, il est important de garder ces relations de cellule et de plage claires.

Voici l’ensemble principal de getters, setters et autres méthodes utiles le plus souvent utilisées dans les scripts. Il s’agit d’un excellent point de départ pour votre parcours d’API. Les sections suivantes groupent les méthodes et aident à créer un modèle mental lorsque vous commencez à déverrouiller les `Range` API de l’objet.

## <a name="example-scripts"></a>Exemples de scripts

* [Lecture et écriture de base](#basic-read-and-write)
* [Ajouter une ligne à la fin de la feuille de calcul](#add-row-at-the-end-of-worksheet)
* [Effacer le filtre de colonne](clear-table-filter-for-active-cell.md)
* [Colorier chaque cellule avec une couleur unique](#color-each-cell-with-unique-color)
* [Mettre à jour la plage avec des valeurs à l’aide d’un tableau 2D](#update-range-with-values-using-2d-array)

### <a name="basic-read-and-write"></a>Lecture et écriture de base

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

### <a name="add-row-at-the-end-of-worksheet"></a>Ajouter une ligne à la fin de la feuille de calcul

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

### <a name="color-each-cell-with-unique-color"></a>Colorier chaque cellule avec une couleur unique

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

### <a name="update-range-with-values-using-2d-array"></a>Mettre à jour la plage avec des valeurs à l’aide d’un tableau 2D

Calcule dynamiquement la dimension de plage à mettre à jour en fonction des valeurs de tableau 2D.

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

## <a name="training-videos-range-basics"></a>Vidéos de formation : informations de base sur les plages

_Informations de base sur les plages_

[![Regarder une vidéo pas à pas sur les principes de base de la plage](../../images/rangebasics-vid.png)](https://youtu.be/4emjkOFdLBA "Vidéo pas à pas sur les principes de base de la plage")

_Ajouter une ligne à la fin de la feuille de calcul_

[![Regardez une vidéo détaillée sur l’ajout d’une ligne à la fin d’une feuille de calcul](../../images/rangebasics-addrow-vid.png)](https://youtu.be/RgtUar013D0 "Vidéo pas à pas sur l’ajout d’une ligne à la fin d’une feuille de calcul")

## <a name="methods-that-return-some-range-metadata"></a>Méthodes qui retournent des métadonnées de plage

* getAddress(), getAddressLocal()
* getCellCount()
* getRowCount(), getColumnCount()

## <a name="methods-that-return-dataconstants-associated-with-a-given-range"></a>Méthodes qui retournent des données/constantes associées à une plage donnée

### <a name="returned-as-single-cell-value"></a>Renvoyé en tant que valeur de cellule unique

* getFormula(), getFormulaLocal()
* getFormulaR1C1()
* getNumberFormat(), getNumberFormatLocal()
* getText()
* getValue()
* getValueType()

### <a name="returned-as-2d-arrays-whole-range"></a>Renvoyé en tant que tableaux 2D (plage entière)

* getFormulas(), getFormulasLocal()
* getFormulasR1C1()
* getNumberFormatCategories()
* getNumberFormats(), getNumberFormatsLocal()
* getTexts()
* getValues()
* getValueTypes()
* getHidden()
* getIsEntireRow()
* getIsEntireColumn()

## <a name="methods-that-return-other-range-object"></a>Méthodes qui retournent un autre objet de plage

* getSurroundingRegion() : similaire à CurrentRegion dans VBA
* getCell(ligne, colonne)
* getColumn(column)
* getColumnHidden()
* getColumnsAfter(count)
* getColumnsBefore(count)
* getEntireColumn()
* getEntireRow()
* getLastCell()
* getLastColumn()
* getLastRow()
* getRow(row)
* getRowHidden()
* getRowsAbove(count)
* getRowsBelow(count)

**Important/Intéressant**

* _workbook_.getSelectedRange()
* _workbook_.getActiveCell()
* getUsedRange(valuesOnly)
* getAbsoluteResizedRange(numRows, numColumns)
* getOffsetRange(rowOffset, columnOffset)
* getResizedRange(deltaRows, deltaColumns)

## <a name="methods-that-return-a-range-object-in-relation-to-another-range-object"></a>Méthodes qui retournent un objet de plage par rapport à un autre objet de plage

* getBoundingRect(anotherRange)
* getIntersection(anotherRange)

## <a name="methods-that-return-other-objects-non-range-objects"></a>Méthodes qui retournent d’autres objets (objets autres que des plages)

* getDirectPrecedents()
* getWorksheet()
* getTables(fullyContained)
* getPivotTables(fullyContained)
* getDataValidation()
* getPredefinedCellStyle()

## <a name="set-methods"></a>Définir des méthodes

### <a name="singular-cell-set-methods"></a>Méthodes de jeu de cellules au singular

* setFormula(formula)
* setFormulaLocal(formulaLocal)
* setFormulaR1C1(formulaR1C1)
* setNumberFormatLocal(numberFormatLocal)
* setValue(value)

### <a name="2d--entire-range-set-methods"></a>2D / méthodes d’ensemble de plages entières

* setFormulas(formulas)
* setFormulasLocal(formulasLocal)
* setFormulasR1C1(formulasR1C1)
* setNumberFormat(numberFormat)
* setNumberFormats(numberFormats)
* setNumberFormatsLocal(numberFormatsLocal)
* setValues(values)

## <a name="other-methods"></a>Autres méthodes

* merge(across)
* unmerge()

## <a name="coming-soon"></a>Bientôt disponible

* API de bordure de plage
