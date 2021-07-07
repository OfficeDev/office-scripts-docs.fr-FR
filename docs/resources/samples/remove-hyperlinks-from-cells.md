---
title: Supprimer des liens hypertexte de chaque cellule d’une Excel de calcul
description: Découvrez comment utiliser des scripts Office pour supprimer des liens hypertexte de chaque cellule d’une Excel de calcul.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: dc33eb639edac8ada29824a53440031942e59179
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313749"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="72b80-103">Supprimer des liens hypertexte de chaque cellule d’une Excel de calcul</span><span class="sxs-lookup"><span data-stu-id="72b80-103">Remove hyperlinks from each cell in an Excel worksheet</span></span>

 <span data-ttu-id="72b80-104">Cet exemple permet d’effacer tous les liens hypertexte de la feuille de calcul actuelle.</span><span class="sxs-lookup"><span data-stu-id="72b80-104">This sample clears all of the hyperlinks from the current worksheet.</span></span> <span data-ttu-id="72b80-105">Il parcourt la feuille de calcul et s’il existe un lien hypertexte associé à la cellule, il effacera le lien hypertexte tout en conservant la valeur de la cellule telle quelle.</span><span class="sxs-lookup"><span data-stu-id="72b80-105">It traverses the worksheet and if there is any hyperlink associated with the cell, it clears the hyperlink yet retains the cell value as is.</span></span> <span data-ttu-id="72b80-106">Enregistre également le temps qu’il faut pour effectuer la traversée.</span><span class="sxs-lookup"><span data-stu-id="72b80-106">Also logs the time it takes to complete traversal.</span></span>

> [!NOTE]
> <span data-ttu-id="72b80-107">Cela fonctionne uniquement si le nombre de cellules est < 10 000.</span><span class="sxs-lookup"><span data-stu-id="72b80-107">This only works if the cell count is < 10k.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="72b80-108">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="72b80-108">Sample Excel file</span></span>

<span data-ttu-id="72b80-109">Téléchargez le <a href="remove-hyperlinks.xlsx"> fichierremove-hyperlinks.xlsx</a> pour un classez prêt à l’emploi.</span><span class="sxs-lookup"><span data-stu-id="72b80-109">Download the file <a href="remove-hyperlinks.xlsx">remove-hyperlinks.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="72b80-110">Ajoutez le script suivant pour essayer l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="72b80-110">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-remove-hyperlinks"></a><span data-ttu-id="72b80-111">Exemple de code : supprimer des liens hypertexte</span><span class="sxs-lookup"><span data-stu-id="72b80-111">Sample code: Remove hyperlinks</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, sheetName: string = 'Sheet1') {
  // Get the active worksheet. 
  let sheet = workbook.getWorksheet(sheetName);

  // Get the used range to operate on.
  // For large ranges (over 10000 entries), consider splitting the operation into batches for performance.
  const targetRange = sheet.getUsedRange(true);
  console.log(`Target Range to clear hyperlinks from: ${targetRange.getAddress()}`);

  const rowCount = targetRange.getRowCount();
  const colCount = targetRange.getColumnCount();
  console.log(`Searching for hyperlinks in ${targetRange.getAddress()} which contains ${(rowCount * colCount)} cells`);

  // Go through each individual cell looking for a hyperlink. 
  // This allows us to limit the formatting changes to only the cells with hyperlink formatting.
  let clearedCount = 0;
  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
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

  console.log(`Done. Cleared hyperlinks from ${clearedCount} cells`);
}
```

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a><span data-ttu-id="72b80-112">Vidéo de formation : supprimer des liens hypertexte de chaque cellule d’une Excel de calcul</span><span class="sxs-lookup"><span data-stu-id="72b80-112">Training video: Remove hyperlinks from each cell in an Excel worksheet</span></span>

<span data-ttu-id="72b80-113">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/v20fdinxpHU).</span><span class="sxs-lookup"><span data-stu-id="72b80-113">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/v20fdinxpHU).</span></span>
