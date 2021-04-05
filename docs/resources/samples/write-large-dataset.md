---
title: Optimisation des performances lors de l’écriture d’un jeu de données de grande taille
description: Découvrez comment optimiser les performances lors de l’écriture d’un jeu de données de grande taille dans les scripts Office.
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: 190072e58238be95a2939f73dcda077ed91db848
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571260"
---
# <a name="performance-optimization-when-writing-a-large-dataset"></a><span data-ttu-id="b0601-103">Optimisation des performances lors de l’écriture d’un jeu de données de grande taille</span><span class="sxs-lookup"><span data-stu-id="b0601-103">Performance optimization when writing a large dataset</span></span>

## <a name="basic-performance-optimization"></a><span data-ttu-id="b0601-104">Optimisation des performances de base</span><span class="sxs-lookup"><span data-stu-id="b0601-104">Basic performance optimization</span></span>

<span data-ttu-id="b0601-105">Pour obtenir des informations de base sur les performances dans Les scripts Office, consultez la [section sur](getting-started.md#basic-performance-considerations) les performances de l’article de mise en place.</span><span class="sxs-lookup"><span data-stu-id="b0601-105">For performance basics in Office Scripts, see the [performance section](getting-started.md#basic-performance-considerations) of the Getting Started article.</span></span>

## <a name="sample-code-optimize-performance-of-a-large-dataset"></a><span data-ttu-id="b0601-106">Exemple de code : optimiser les performances d’un jeu de données de grande taille</span><span class="sxs-lookup"><span data-stu-id="b0601-106">Sample code: Optimize performance of a large dataset</span></span>

<span data-ttu-id="b0601-107">`setValues()`L’API Range permet de définir les valeurs d’une plage.</span><span class="sxs-lookup"><span data-stu-id="b0601-107">The `setValues()` Range API allows setting the values of a range.</span></span> <span data-ttu-id="b0601-108">Cette API présente des limitations de données en fonction de divers facteurs tels que la taille des données, les paramètres réseau, etc. Pour mettre à jour de manière fiable une grande plage de données, vous devez penser à des mises à jour de données en plus petits blocs.</span><span class="sxs-lookup"><span data-stu-id="b0601-108">This API has data limitations depending on various factors such as data size, network settings, etc. In order to reliably update a large range of data, you'll need to think about doing data updates in smaller chunks.</span></span> <span data-ttu-id="b0601-109">Ce script tente de le faire et écrit les lignes d’une plage en blocs afin que si une grande plage doit être mise à jour, elle peut être effectuée dans des parties plus petites.</span><span class="sxs-lookup"><span data-stu-id="b0601-109">This script attempts to do this and writes rows of a range in chunks so that if a large range needs to be updated, it can be done in smaller parts.</span></span> <span data-ttu-id="b0601-110">**Avertissement**: il n’a pas été testé dans différentes tailles, donc n’ignorez pas cela si vous souhaitez l’utiliser dans votre script.</span><span class="sxs-lookup"><span data-stu-id="b0601-110">**Warning**: It has not been tested across various sizes so be aware of that if you want to use this in your script.</span></span> <span data-ttu-id="b0601-111">Comme nous avons la possibilité de tester, nous allons mettre à jour les résultats concernant son fonctionnement pour différentes tailles de données.</span><span class="sxs-lookup"><span data-stu-id="b0601-111">As we have opportunity to test, we'll update with findings around how it performs for various data sizes.</span></span>

<span data-ttu-id="b0601-112">Ce script sélectionne 1 000 cellules par bloc, mais vous pouvez le remplacer pour tester son fonctionnement.</span><span class="sxs-lookup"><span data-stu-id="b0601-112">This script selects 1K cells per chunk but you can override to test out how it works for you.</span></span> <span data-ttu-id="b0601-113">Il met à jour 100 000 lignes avec 6 colonnes de données.</span><span class="sxs-lookup"><span data-stu-id="b0601-113">It updates 100k rows with 6 columns of data.</span></span> <span data-ttu-id="b0601-114">Exécutez cette information sur une feuille vierge pour l’examiner.</span><span class="sxs-lookup"><span data-stu-id="b0601-114">Run this on a blank sheet to examine.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  let data: (string | number | boolean)[][] = [];
  // Number of rows in the random data (x 6 columns).
  const sampleRows = 100000;

  console.log(`Generating data...`)
  // Dynamically generate some random data for testing purpose. 
  for (let i = 0; i < sampleRows; i++) {
    data.push([i, ...[getRandomString(5), getRandomString(20), getRandomString(10), Math.random()], "Sample data"]);
  }

  console.log(`Calling update range function...`);
  const updated = updateRangeInChunks(sheet.getRange("B2"), data);
  if (!updated) {
    console.log(`Update did not take place or complete. Check and run again.`)
  }

  return;
}

function updateRangeInChunks(
  startCell: ExcelScript.Range,
  values: (string | boolean | number)[][],
  cellsInChunk: number = 10000
): boolean {

  const startTime = new Date().getTime();
  console.log(`Cells per chunk setting: ${cellsInChunk}`);
  if (!values) {
    console.log(`Invalid input values to update.`);
    return false;
  }
  if (values.length === 0 || values[0].length === 0) {
    console.log(`Empty data -- nothing to update.`);
    return true;
  }
  const totalCells = values.length * values[0].length;

  console.log(`Total cells to update in the target range: ${totalCells}`);
  if (totalCells <= cellsInChunk) {
    console.log(`No need to chunk -- updating directly`);
    updateTargetRange(startCell, values);
    return true;
  }

  const rowsPerChunk = Math.floor(cellsInChunk / values[0].length);
  console.log("Rows per chunk: " + rowsPerChunk);
  let rowCount = 0;
  let totalRowsUpdated = 0;
  let chunkCount = 0;

  for (let i = 0; i < values.length; i++) {
    rowCount++;
    if (rowCount === rowsPerChunk) {
      chunkCount++;
      console.log(`Calling update next chunk function. Chunk#: ${chunkCount}`);
      updateNextChunk(startCell, values, rowsPerChunk, totalRowsUpdated);
      rowCount = 0;
      totalRowsUpdated += rowsPerChunk;
      console.log(`${((totalRowsUpdated / values.length) * 100).toFixed(1)}% Done`);

    }
  }
  console.log(`Updating remaining rows -- last chunk: ${rowCount}`)
  if (rowCount > 0) {
    updateNextChunk(startCell, values, rowCount, totalRowsUpdated);
  }

  let endTime = new Date().getTime();
  console.log(`Completed ${totalCells} cells update. It took: ${((endTime - startTime) / 1000).toFixed(6)} seconds to complete. ${((((endTime  - startTime) / 1000)) / cellsInChunk).toFixed(8)} seconds per ${cellsInChunk} cells-chunk.`);

  return true;
}

/**
 * A helper function that computes the target range and updates. 
 */

function updateNextChunk(
  startingCell: ExcelScript.Range,
  data: (string | boolean | number)[][],
  rowsPerChunk: number,
  totalRowsUpdated: number
) {

  const newStartCell = startingCell.getOffsetRange(totalRowsUpdated, 0);
  const targetRange = newStartCell.getResizedRange(rowsPerChunk - 1, data[0].length - 1);
  console.log(`Updating chunk at range ${targetRange.getAddress()}`);
  const dataToUpdate = data.slice(totalRowsUpdated, totalRowsUpdated + rowsPerChunk);
  try {
    targetRange.setValues(dataToUpdate);
  } catch (e) {
    throw `Error while updating the chunk range: ${JSON.stringify(e)}`;
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

## <a name="training-video-optimize-performance-when-writing-a-large-dataset"></a><span data-ttu-id="b0601-115">Vidéo de formation : optimiser les performances lors de l’écriture d’un jeu de données de grande taille</span><span class="sxs-lookup"><span data-stu-id="b0601-115">Training video: Optimize performance when writing a large dataset</span></span>

<span data-ttu-id="b0601-116">[![Regardez une vidéo sur l’optimisation des performances lors de l’écriture d’un jeu de données de grande taille](../../images/largedata-vid.png)](https://youtu.be/BP9Kp0Ltj7U "Vidéo sur l’optimisation des performances lors de l’écriture d’un jeu de données de grande taille")</span><span class="sxs-lookup"><span data-stu-id="b0601-116">[![Watch video on how to optimize performance when writing a large dataset](../../images/largedata-vid.png)](https://youtu.be/BP9Kp0Ltj7U "Video on how to optimize performance when writing a large dataset")</span></span>