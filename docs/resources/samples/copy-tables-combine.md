---
title: Combiner les données de plusieurs tables Excel dans une seule table
description: Découvrez comment utiliser Office scripts pour combiner les données de plusieurs tables Excel dans une seule table.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: f8285ac2302ef30de66c2bdbfb9277f83a3a75690d59a1e9cb066f544eeffb19
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846973"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>Combiner les données de plusieurs tables Excel dans une seule table

Cet exemple combine les données de plusieurs tables Excel dans une seule table qui inclut toutes les lignes. Il suppose que toutes les tables utilisées ont la même structure.

Il existe deux variantes de ce script :

1. Le [premier script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combine toutes les tables du Excel fichier.
1. Le [deuxième script obtient](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) de manière sélective les tableaux d’un ensemble de feuilles de calcul.

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez <a href="tables-copy.xlsx">tables-copy.xlsx</a> pour un livre de travail prêt à l’emploi. Ajoutez les scripts suivants pour essayer l’exemple vous-même !

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Exemple de code : combiner les données de plusieurs tables Excel dans une seule table

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');
  
  // Get the header values for the first table in the workbook.
  // This also saves the table list before we add the new, combined table.
  const tables = workbook.getTables();    
  const headerValues = tables[0].getHeaderRowRange().getTexts();
  console.log(headerValues);

  // Copy the headers on a new worksheet to an equal-sized range.
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);

  // Add the data from each table in the workbook to the new table.
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
  for (let table of tables) {      
    let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
    let rowCount = table.getRowCount();

    // If the table is not empty, add its rows to the combined table.
    if (rowCount > 0) {
      combinedTable.addRows(-1, dataValues);
    }
  }
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>Exemple de code : combiner les données de plusieurs tables Excel dans des feuilles de calcul sélectionnées dans un seul tableau

Téléchargez l’exemple <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> fichier et utilisez-le avec le script suivant pour l’essayer vous-même !

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Set the worksheet names to get tables from.
  const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');

  // Create a new table with the same headers as the other tables.
  const headerValues = workbook.getWorksheet(sheetNames[0]).getTables()[0].getHeaderRowRange().getTexts();
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);

  // Go through each listed worksheet and get their tables.
  sheetNames.forEach((sheet) => {
    const tables = workbook.getWorksheet(sheet).getTables();     
    for (let table of tables) {
      // Get the rows from the tables.
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();

      // If there's data in the table, add it to the combined table.
      if (rowCount > 0) {
          combinedTable.addRows(-1, dataValues);
      }
    }
  });
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>Vidéo de formation : Combiner les données de plusieurs tables Excel dans une seule table

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/di-8JukK3Lc).
