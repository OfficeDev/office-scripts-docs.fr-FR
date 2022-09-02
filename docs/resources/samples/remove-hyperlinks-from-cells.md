---
title: Supprimer des liens hypertexte de chaque cellule d’une feuille de calcul Excel
description: Découvrez comment utiliser les scripts Office pour supprimer des liens hypertexte de chaque cellule d’une feuille de calcul Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1445988b1e6a85fcab8914ffeaaef80a07a52f5e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572625"
---
# <a name="remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Supprimer des liens hypertexte de chaque cellule d’une feuille de calcul Excel

 Cet exemple efface tous les liens hypertexte de la feuille de calcul active. Il traverse la feuille de calcul et, s’il existe un lien hypertexte associé à la cellule, il efface le lien hypertexte tout en conservant la valeur de la cellule en l’état. Enregistre également le temps nécessaire à la traversée.

> [!NOTE]
> Cela ne fonctionne que si le nombre de cellules est < 10 000.

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez le [ fichierremove-hyperlinks.xlsx](remove-hyperlinks.xlsx) pour un classeur prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-remove-hyperlinks"></a>Exemple de code : Supprimer des liens hypertexte

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

## <a name="training-video-remove-hyperlinks-from-each-cell-in-an-excel-worksheet"></a>Vidéo de formation : Supprimer des liens hypertexte de chaque cellule d’une feuille de calcul Excel

[Regardez Sudhi Ramamurthy parcourir cet exemple sur YouTube](https://youtu.be/v20fdinxpHU).
