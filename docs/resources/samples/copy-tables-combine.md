---
title: Combiner les données de plusieurs tables Excel dans une seule table
description: Découvrez comment utiliser Office scripts pour combiner les données de plusieurs tables Excel dans une seule table.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: ac8c7d0a3f0f4f3d7d3217ffac31aff1a5595d17
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232444"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="80592-103">Combiner les données de plusieurs tables Excel dans une seule table</span><span class="sxs-lookup"><span data-stu-id="80592-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="80592-104">Cet exemple combine les données de plusieurs tables Excel dans une seule table qui inclut toutes les lignes.</span><span class="sxs-lookup"><span data-stu-id="80592-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="80592-105">Il suppose que toutes les tables utilisées ont la même structure.</span><span class="sxs-lookup"><span data-stu-id="80592-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="80592-106">Il existe deux variantes de ce script :</span><span class="sxs-lookup"><span data-stu-id="80592-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="80592-107">Le [premier script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combine toutes les tables du Excel fichier.</span><span class="sxs-lookup"><span data-stu-id="80592-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="80592-108">Le [deuxième script obtient](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) de manière sélective les tableaux d’un ensemble de feuilles de calcul.</span><span class="sxs-lookup"><span data-stu-id="80592-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="80592-109">Exemple de code : combiner les données de plusieurs tables Excel dans une seule table</span><span class="sxs-lookup"><span data-stu-id="80592-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="80592-110">Téléchargez l’exemple <a href="tables-copy.xlsx">tables-copy.xlsx</a> fichier et utilisez-le avec le script suivant pour l’essayer vous-même !</span><span class="sxs-lookup"><span data-stu-id="80592-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="80592-111">Exemple de code : combiner les données de plusieurs tables Excel dans des feuilles de calcul sélectionnées dans un seul tableau</span><span class="sxs-lookup"><span data-stu-id="80592-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="80592-112">Téléchargez l’exemple <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> fichier et utilisez-le avec le script suivant pour l’essayer vous-même !</span><span class="sxs-lookup"><span data-stu-id="80592-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="80592-113">Vidéo de formation : Combiner les données de plusieurs tables Excel dans une seule table</span><span class="sxs-lookup"><span data-stu-id="80592-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="80592-114">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/di-8JukK3Lc).</span><span class="sxs-lookup"><span data-stu-id="80592-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/di-8JukK3Lc).</span></span>
