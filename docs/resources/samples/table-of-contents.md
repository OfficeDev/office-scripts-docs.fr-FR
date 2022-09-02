---
title: Créer une table des matières de classeur
description: Découvrez comment créer une table des matières avec des liens vers chaque feuille de calcul.
ms.date: 01/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5b158160ecb9ac29df547c6da6552e21c9875be3
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572513"
---
# <a name="create-a-workbook-table-of-contents"></a>Créer une table des matières de classeur

Cet exemple montre comment créer une table des matières pour le classeur. Chaque entrée de la table des matières est un lien hypertexte vers l’une des feuilles de calcul du classeur.

:::image type="content" source="../../images/table-of-contents-sample.png" alt-text="Feuille de calcul de table des matières affichant des liens vers les autres feuilles de calcul.":::

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez [table-of-contents.xlsx](table-of-contents.xlsx) pour un classeur prêt à l’emploi. Ajoutez le script suivant et essayez l’exemple vous-même !

## <a name="sample-code-create-a-workbook-table-of-contents"></a>Exemple de code : Créer une table des matières d’un classeur

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
