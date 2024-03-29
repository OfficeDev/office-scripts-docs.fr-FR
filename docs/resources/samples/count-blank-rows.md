---
title: Compter les lignes vides sur les feuilles
description: Découvrez comment utiliser des scripts Office pour détecter s’il existe des lignes vides au lieu de données dans des feuilles de calcul, puis signaler le nombre de lignes vierges à utiliser dans un flux Power Automate données.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1ae513928b885994dc7f6d1b8ad66d694b61e7b7
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585561"
---
# <a name="count-blank-rows-on-sheets"></a>Compter les lignes vides sur les feuilles

Ce projet comprend deux scripts :

* [Compter les lignes vides sur une feuille](#sample-code-count-blank-rows-on-a-given-sheet) donnée : parcourt la plage utilisée dans une feuille de calcul donnée et renvoie un nombre de lignes vide.
* [Compter les lignes vides sur toutes les feuilles](#sample-code-count-blank-rows-on-all-sheets) : parcourt la plage utilisée sur toutes les _feuilles_ de calcul et renvoie un nombre de lignes vide.

> [!NOTE]
> Pour notre script, une ligne vide est toute ligne sans données. La ligne peut avoir une mise en forme.

_Cette feuille renvoie le nombre de 4 lignes vides_

:::image type="content" source="../../images/blank-rows.png" alt-text="Feuille de calcul affichant des données avec des lignes vides.":::

_Cette feuille renvoie le nombre de 0 lignes vides (toutes les lignes ont des données)_

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Feuille de calcul montrant les données sans lignes vides.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a>Exemple de code : compter les lignes vides sur une feuille donnée

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Get the worksheet named "Sheet1".
  const sheet = workbook.getWorksheet('Sheet1'); 
  
  // Get the entire data range.
  const range = sheet.getUsedRange(true);

  // If the used range is empty, end the script.
  if (!range) {
    console.log(`No data on this sheet.`);
    return;
  }
  
  // Log the address of the used range.
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
  // Look through the values in the range for blank rows.
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let emptyRow = true;
    
    // Look at every cell in the row for one with a value.
    for (let cell of row) {
      if (cell.toString().length > 0) {
        emptyRow = false
      }
    }

    // If no cell had a value, the row is empty.
    if (emptyRow) {
      emptyRows++;
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a>Exemple de code : compter les lignes vides sur toutes les feuilles

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  // Loop through every worksheet in the workbook.
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) {     
    // Get the entire data range.
    const range = sheet.getUsedRange(true);
  
    // If the used range is empty, skip to the next worksheet.
    if (!range) {
      console.log(`No data on this sheet.`);
      continue;
    }
    
    // Log the address of the used range.
    console.log(`Used range for the worksheet: ${range.getAddress()}`);
      
    // Look through the values in the range for blank rows.
    const values = range.getValues();
    for (let row of values) {
      let emptyRow = true;
      
      // Look at every cell in the row for one with a value.
      for (let cell of row) {
        if (cell.toString().length > 0) {
          emptyRow = false
        }
      }
  
      // If no cell had a value, the row is empty.
      if (emptyRow) {
        emptyRows++;
      }
    }
  }

  // Log the number of empty rows.
  console.log(`Total empty rows: ${emptyRows}`);

  // Return the number of empty rows for use in a Power Automate flow.
  return emptyRows;
}
```
