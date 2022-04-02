---
title: Créer une table des matières de classeur
description: Découvrez comment créer une table des matières avec des liens vers chaque feuille de calcul.
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: b2d69609514c2e1e87f9c0590ea10152fc7d5e7d
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585519"
---
# <a name="create-a-workbook-table-of-contents"></a>Créer une table des matières de classeur

Cet exemple montre comment créer une table des matières pour le workbook. Chaque entrée de la table des matières est un lien hypertexte vers l’une des feuilles de calcul du manuel.

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="Feuille de calcul table des matières affichant des liens vers les autres feuilles de calcul.":::

## <a name="sample-excel-file"></a>Exemple Excel fichier

<a href="table-of-contents.xlsx"> Téléchargeztable-of-contents.xlsx</a> pour un livre de travail prêt à l’emploi. Ajoutez le script suivant et essayez l’exemple vous-même !

## <a name="sample-code-create-a-workbook-table-of-contents"></a>Exemple de code : créer une table des matières de workbook

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Insert a new worksheet at the beginning of the workbook.
  let tocSheet = workbook.addWorksheet();
  tocSheet.setPosition(0);
  tocSheet.setName("Table of Contents");

  // Give the worksheet a title in the sheet.
  tocSheet.getRange("A1").setValue("Table of Contents");
  tocSheet.getRange("A1").getFormat().getFont().setBold(true);

  // Create the table of contents headers.
  let tocRange = tocSheet.getRange("A2:B2")
  tocRange.setValues([["#", "Name"]]);

  // Get the range for the table of contents entries.
  let worksheets = workbook.getWorksheets();
  tocRange = tocRange.getResizedRange(worksheets.length, 0);

  // Loop through all worksheets in the workbook, except the first one.
  for (let i = 1; i < worksheets.length; i++) {
    // Create a row for each worksheet with its index and linked name.
    tocRange.getCell(i, 0).setValue(i);
    tocRange.getCell(i, 1).setHyperlink({
      textToDisplay: worksheets[i].getName(),
      documentReference: `'${worksheets[i].getName()}'!A1`
    });
  };

  // Activate the table of contents worksheet.
  tocSheet.activate();
}
```
