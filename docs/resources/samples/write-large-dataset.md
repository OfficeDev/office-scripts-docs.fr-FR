---
title: Écrire un jeu de données de grande taille
description: Découvrez comment fractionner un jeu de données de grande taille en opérations d’écriture plus petites Office scripts.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 06abb58c61c18620d638ab3eb61ea68398bf20aa
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545622"
---
# <a name="write-a-large-dataset"></a><span data-ttu-id="87536-103">Écrire un jeu de données de grande taille</span><span class="sxs-lookup"><span data-stu-id="87536-103">Write a large dataset</span></span>

<span data-ttu-id="87536-104">`Range.setValues()`L’API place les données dans une plage.</span><span class="sxs-lookup"><span data-stu-id="87536-104">The `Range.setValues()` API puts data in a range.</span></span> <span data-ttu-id="87536-105">Cette API présente des limitations en fonction de différents facteurs, tels que la taille des données et les paramètres réseau.</span><span class="sxs-lookup"><span data-stu-id="87536-105">This API has limitations depending on various factors, such as data size and network settings.</span></span> <span data-ttu-id="87536-106">Cela signifie que si vous essayez d’écrire une grande quantité d’informations dans un workbook en une seule opération, vous devrez écrire les données par lots plus petits afin de mettre à jour de manière fiable une grande plage [.](../../testing/platform-limits.md)</span><span class="sxs-lookup"><span data-stu-id="87536-106">This means that if you attempt to write a massive amount of information to a workbook as a single operation, you'll need to write the data in smaller batches in order to reliably update a [large range](../../testing/platform-limits.md).</span></span>

<span data-ttu-id="87536-107">Pour obtenir des informations de base sur les performances Office scripts, veuillez lire Améliorer les performances de [vos scripts Office.](../../develop/web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="87536-107">For performance basics in Office Scripts, please read [Improve the performance of your Office Scripts](../../develop/web-client-performance.md).</span></span>

## <a name="sample-code-write-a-large-dataset"></a><span data-ttu-id="87536-108">Exemple de code : écrire un jeu de données de grande taille</span><span class="sxs-lookup"><span data-stu-id="87536-108">Sample code: Write a large dataset</span></span>

<span data-ttu-id="87536-109">Ce script écrit les lignes d’une plage dans des parties plus petites.</span><span class="sxs-lookup"><span data-stu-id="87536-109">This script writes rows of a range in smaller parts.</span></span> <span data-ttu-id="87536-110">Il sélectionne 1 000 cellules à écrire à la fois.</span><span class="sxs-lookup"><span data-stu-id="87536-110">It selects 1000 cells to write at a time.</span></span> <span data-ttu-id="87536-111">Exécutez le script sur une feuille de calcul vide pour voir les lots de mise à jour en action.</span><span class="sxs-lookup"><span data-stu-id="87536-111">Run the script on a blank worksheet to see the update batches in action.</span></span> <span data-ttu-id="87536-112">La sortie de la console fournit des informations supplémentaires sur ce qui se passe.</span><span class="sxs-lookup"><span data-stu-id="87536-112">The console output gives further insight into what's happening.</span></span>

> [!NOTE]
> <span data-ttu-id="87536-113">Vous pouvez modifier le nombre total de lignes écrites en modifiant la valeur de `SAMPLE_ROWS` .</span><span class="sxs-lookup"><span data-stu-id="87536-113">You can change the number of total rows being written by changing the value of `SAMPLE_ROWS`.</span></span> <span data-ttu-id="87536-114">Vous pouvez modifier le nombre de cellules à écrire en tant qu’action unique en modifiant la valeur de `CELLS_IN_BATCH` .</span><span class="sxs-lookup"><span data-stu-id="87536-114">You can change the number of cells to write as a single action by changing the value of `CELLS_IN_BATCH`.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const SAMPLE_ROWS = 100000;
  const CELLS_IN_BATCH = 10000;

  // Get the current worksheet.
  const sheet = workbook.getActiveWorksheet();

  console.log(`Generating data...`)
  let data: (string | number | boolean)[][] = [];
  // Generate six columns of random data per row. 
  for (let i = 0; i < SAMPLE_ROWS; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInBatches(sheet.getRange("B2"), data, CELLS_IN_BATCH);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`);
  }
}

function updateRangeInBatches(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInBatch: number
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per batch setting: ${cellsInBatch}`);

  // Determine the total number of cells to write.
  const totalCells = values.length * values[0].length;
  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInBatch) {
    console.log(`No need to batch -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  // Determine how many rows to write at once.
  const rowsPerBatch = Math.floor(cellsInBatch / values[0].length);
  console.log("Rows per batch: " + rowsPerBatch);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let batchCount = 0;

  // Write each batch of rows.
  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerBatch) {
      batchCount++;
      console.log(`Calling update next batch function. Batch#: ${batchCount}`);
      updateNextBatch(startCell, values, rowsPerBatch, totalRowsUpdated);

      // Write a completion percentage to help the user understand the progress.
      rowCount = 0;
      totalRowsUpdated += rowsPerBatch;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);
    }
  }
  
  console.log(`Updating remaining rows -- last batch: ${rowCount}`)
  if (rowCount > 0) {
    updateNextBatch(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInBatch).toFixed(8)} seconds per ${cellsInBatch} cells-batch.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */
function updateNextBatch(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerBatch: number,
  totalRowsUpdated: number
) {
  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerBatch - 1, data[0].length - 1);
  console.log(`Updating batch at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerBatch);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the batch range: ${JSON.stringify(e)}`;
  }
  return;
}

/**
 * A helper function that computes the target range given the target range's starting cell
 * and selected range and updates the values.
 */
function updateTargetRange(
  targetCell: ExcelScript.Range,
  values: (string | boolean | number)[][]
) {
  const targetRange = targetCell.getResizedRange(values.length - 1, values[0].length - 1);
  console.log(`Updating the range: ${targetRange.getAddress()}`);
  try {
    targetRange.setValues(values);
  } catch (e) {
    throw `Error while updating the whole range: ${JSON.stringify(e)}`;
  }
  return;
}

// Credit: https://www.codegrepper.com/code-examples/javascript/random+text+generator+javascript
function getRandomString(length: number): string {
  var randomChars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < length; i++) {
    result += randomChars.charAt(Math.floor(Math.random() * randomChars.length));
  }
  return result;
}
```

## <a name="training-video-write-a-large-dataset"></a><span data-ttu-id="87536-115">Vidéo de formation : Écrire un jeu de données de grande taille</span><span class="sxs-lookup"><span data-stu-id="87536-115">Training video: Write a large dataset</span></span>

<span data-ttu-id="87536-116">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/BP9Kp0Ltj7U).</span><span class="sxs-lookup"><span data-stu-id="87536-116">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/BP9Kp0Ltj7U).</span></span>
