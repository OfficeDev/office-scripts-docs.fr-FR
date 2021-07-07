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
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Supprimer des liens hypertexte de chaque cellule d’une Excel de calcul

 Cet exemple permet d’effacer tous les liens hypertexte de la feuille de calcul actuelle. Il parcourt la feuille de calcul et s’il existe un lien hypertexte associé à la cellule, il effacera le lien hypertexte tout en conservant la valeur de la cellule telle quelle. Enregistre également le temps qu’il faut pour effectuer la traversée.

> [!NOTE]
> Cela fonctionne uniquement si le nombre de cellules est < 10 000.

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez le <a href="remove-hyperlinks.xlsx"> fichierremove-hyperlinks.xlsx</a> pour un classez prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-remove-hyperlinks"></a>Exemple de code : supprimer des liens hypertexte

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Vidéo de formation : supprimer des liens hypertexte de chaque cellule d’une Excel de calcul

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/v20fdinxpHU).
