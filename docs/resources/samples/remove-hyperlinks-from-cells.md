---
title: Supprimer des liens hypertexte de chaque cellule d’une Excel de calcul
description: Découvrez comment utiliser des scripts Office pour supprimer des liens hypertexte de chaque cellule d’une Excel de calcul.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: eb5f486cb5228e639727c5ee7e6c335d5e94239f
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232745"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="43466-103">Supprimer des liens hypertexte de chaque cellule d’une Excel de calcul</span><span class="sxs-lookup"><span data-stu-id="43466-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="43466-104">Cet exemple permet d’effacer tous les liens hypertexte de la feuille de calcul actuelle.</span><span class="sxs-lookup"><span data-stu-id="43466-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="43466-105">Il parcourt la feuille de calcul et, s’il existe un lien hypertexte associé à la cellule, il effacera le lien hypertexte tout en conservant la valeur de la cellule telle quelle.</span><span class="sxs-lookup"><span data-stu-id="43466-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="43466-106">Enregistre également le temps qu’il faut pour effectuer la traversée.</span><span class="sxs-lookup"><span data-stu-id="43466-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="43466-107">Cela fonctionne uniquement si le nombre de cellules est < 10 000.</span><span class="sxs-lookup"><span data-stu-id="43466-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="43466-108">Exemple de code : supprimer des liens hypertexte</span><span class="sxs-lookup"><span data-stu-id="43466-108">Sample code: Remove hyperlinks</span></span>

<span data-ttu-id="43466-109">Téléchargez le fichier <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> utilisé dans cet exemple et testez-le vous-même !</span><span class="sxs-lookup"><span data-stu-id="43466-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> used in this sample and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {

  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);
  const targetRange = sheet.getUsedRange(true);
  if (!targetRange) {
    console.log(`There is no data in the worksheet. `)
    return;
  }
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  const totalCells = rowCount * colCount;
  if (totalCells > 10000) {
    console.log("Too many cells to operate with. Consider editing script to use selected range and then remove hyperlinks in batches. " + targetRange.getAddress());
    return;
  }
  // Call the helper function to remove the hyperlinks. 
  removeHyperLink(targetRange);
  return;
}

/**
 * Removes hyperlink for each cell in the target range. Logs the time it takes to complete traversal.
 * @param targetRange Target range to clear the hyperlinks from.
 */
function removeHyperLink(targetRange: ExcelScript.Range): void {
  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);
  let clearedCount = 0;
  let cellsVisited = 0;

  let groupStart = new Date().getTime();
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      cellsVisited++;
      if (cellsVisited % 50 === 0) {
        let groupEnd = new Date().getTime();
        console.log(`Completed ${cellsVisited} cells out of ${rowCount * colCount}. This group took: ${(groupEnd - groupStart) / 1000} seconds to complete.`);
        groupStart = new Date().getTime();
      }
      const cell = targetRange.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        cell.clear(ExcelScript.ClearApplyTo.hyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Black');
        clearedCount++;
      }
    }
  }
  console.log(`Done. Inspected ${cellsVisited} cells. Cleared hyperlinks in: ${clearedCount} cells`);
  return;
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="43466-110">Vidéo de formation : supprimer des liens hypertexte de chaque cellule d’une Excel de calcul</span><span class="sxs-lookup"><span data-stu-id="43466-110">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="43466-111">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/v20fdinxpHU).</span><span class="sxs-lookup"><span data-stu-id="43466-111">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/v20fdinxpHU).</span></span>
