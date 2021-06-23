---
title: Déplacer des lignes entre des tableaux à l’aide Office Scripts
description: Découvrez comment déplacer des lignes d’une table à l’autre en enregistrement des filtres, puis en traitant et réappliquent les filtres.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: c850ed055457f6733694027469a96a87e74ef66a
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074451"
---
# <a name="move-rows-across-tables-by-saving-filters-then-processing-and-reapplying-the-filters"></a>Déplacer des lignes entre les tables en enregistrer les filtres, puis en traitant et réappliquer les filtres

Ce script effectue les opérations suivantes :

* Sélectionne des lignes dans la table source où la valeur d’une colonne est égale à _une valeur._
* Déplace toutes les lignes sélectionnées dans un autre tableau (cible) d’une autre feuille de calcul.
* Réapplicité des filtres pertinents sur la table source.

:::image type="content" source="../../images/table-filter-before-after.png" alt-text="Captures d’écran du workbook avant et après.":::

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez le fichier <a href="input-table-filters.xlsx">input-table-filters.xlsx</a> utilisé dans cette solution pour l’essayer vous-même !

## <a name="sample-code-move-rows-using-range-values"></a>Exemple de code : déplacer des lignes à l’aide de valeurs de plage

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';
  const IndexOfColumnToFilterOn = 1;
  const NameOfColumnToFilterOn = 'Category';
  const ValueToFilterOn = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TargetTableName}) and target table (${SourceTableName}) are present before running the script. `);
    return;
  }

  // Save the filter criteria.
  const tableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let colFilterCriteria = column.getFilter().getCriteria();
    if (colFilterCriteria) {
      tableFilters[column.getName()] = colFilterCriteria;
    }
  });

  // Get all the data from the table.
  const sourceRange = sourceTable.getRangeBetweenHeaderAndTotal();
  const dataRows: (number | string | boolean)[][] = sourceTable.getRangeBetweenHeaderAndTotal().getValues();

  // Create variables to hold the rows to be moved and their addresses.
  let rowsToMoveValues: (number | string | boolean)[][] = [];
  let rowAddressToRemove: string[] = [];

  // Get the data values from the source table.
  for (let i = 0; i < dataRows.length; i++) { 
    if (dataRows[i][IndexOfColumnToFilterOn] === ValueToFilterOn) {
      rowsToMoveValues.push(dataRows[i]);

      // Get the intersection between table address and the entire row where we found the match. This provides the address of the range to remove.
      let address = sourceRange.getIntersection(sourceRange.getCell(i,0).getEntireRow()).getAddress();
      rowAddressToRemove.push(address);
    }
  }

  // If there are no data rows to process, end the script.
  if (rowsToMoveValues.length < 1) {
    console.log('No rows selected from the source table match the filter criteria.');
    return;
  }

  console.log(`Adding ${rowsToMoveValues.length} rows to target table.`);

  // Insert rows at the end of target table.
  targetTable.addRows(-1, rowsToMoveValues)

  // Remove the rows from the source table.
  const sheet = sourceTable.getWorksheet();

  // Remove all filters before removing rows.
  sourceTable.getAutoFilter().clearCriteria();

  // Important: Remove the rows starting at the bottom of the table.
  // Otherwise, the lower rows change position before they are deleted.
  console.log(`Removing ${rowAddressToRemove.length} rows from the source table.`);
  rowAddressToRemove.reverse().forEach((address) => {
    sheet.getRange(address).delete(ExcelScript.DeleteShiftDirection.up);
  });

  // Reapply the original filters. 
  Object.keys(tableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(tableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a>Vidéo de formation : déplacer des lignes dans des tableaux

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/_3t3Pk4i2L0). Deux scripts sont affichés dans la solution de la vidéo. La principale différence est la façon dont les lignes sont sélectionnées.

* Dans la première variante, les lignes sont sélectionnées en appliquant le filtre de tableau et en lisant la plage visible.
* Dans la seconde, les lignes sont sélectionnées en lisant les valeurs et en extrayant les valeurs de ligne (ce que l’exemple de cette page utilise).
