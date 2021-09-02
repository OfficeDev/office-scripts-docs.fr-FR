---
title: Déplacer des lignes entre des tableaux à l’aide Office Scripts
description: Découvrez comment déplacer des lignes d’une table à l’autre en enregistrement des filtres, puis en traitant et réappliquent les filtres.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 54a41bddaebd4916e8bcffc7bc24f9a219c3d6a4
ms.sourcegitcommit: 6654aeae8a3ee2af84b4d4c4d8ff45b360a303eb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2021
ms.locfileid: "58862074"
---
# <a name="move-rows-across-tables"></a>Déplacer des lignes dans des tableaux

Ce script effectue les opérations suivantes :

* Sélectionne des lignes dans la table source où la valeur d’une colonne est égale à une valeur `FILTER_VALUE` (dans le script).
* Déplace toutes les lignes sélectionnées dans le tableau cible dans une autre feuille de calcul.
* Réapplicité les filtres pertinents dans la table source.

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez le <a href="input-table-filters.xlsx"> fichierinput-table-filters.xlsx</a> pour un classez prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-move-rows-using-range-values"></a>Exemple de code : déplacer des lignes à l’aide de valeurs de plage

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // You can change these names to match the data in your workbook.
  const TARGET_TABLE_NAME = 'Table1';
  const SOURCE_TABLE_NAME = 'Table2';

  // Select what will be moved between tables.
  const FILTER_COLUMN_INDEX = 1;
  const FILTER_VALUE = 'Clothing';

  // Get the Table objects.
  let targetTable = workbook.getTable(TARGET_TABLE_NAME);
  let sourceTable = workbook.getTable(SOURCE_TABLE_NAME);

  // If either table is missing, report that information and stop the script.
  if (!targetTable || !sourceTable) {
    console.log(`Tables missing - Check to make sure both source (${TARGET_TABLE_NAME}) and target table (${SOURCE_TABLE_NAME}) are present before running the script. `);
    return;
  }

  // Save the filter criteria currently on the source table.
  const originalTableFilters = {};
  // For each table column, collect the filter criteria on that column.
  sourceTable.getColumns().forEach((column) => {
    let originalColumnFilter = column.getFilter().getCriteria();
    if (originalColumnFilter) {
      originalTableFilters[column.getName()] = originalColumnFilter;
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
    if (dataRows[i][FILTER_COLUMN_INDEX] === FILTER_VALUE) {
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
  Object.keys(originalTableFilters).forEach((columnName) => {
      sourceTable.getColumnByName(columnName).getFilter().apply(originalTableFilters[columnName]);
    });
}
```

## <a name="training-video-move-rows-across-tables"></a>Vidéo de formation : déplacer des lignes dans des tableaux

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/_3t3Pk4i2L0). Deux scripts sont affichés dans la solution de la vidéo. La principale différence est la façon dont les lignes sont sélectionnées.

* Dans la première variante, les lignes sont sélectionnées en appliquant le filtre de tableau et en lisant la plage visible.
* Dans la seconde, les lignes sont sélectionnées en lisant les valeurs et en extrayant les valeurs de ligne (ce que l’exemple de cette page utilise).
