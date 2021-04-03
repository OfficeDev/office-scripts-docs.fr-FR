---
title: Compter les lignes vides sur les feuilles
description: Découvrez comment utiliser les scripts Office pour détecter s’il existe des lignes vides au lieu de données dans les feuilles de calcul, puis signaler le nombre de lignes vierges à utiliser dans un flux Power Automate.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 1f52b9c4d538d5d3e64dc61dae3e27d046b56862
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571200"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="fccbd-103">Compter les lignes vides sur les feuilles</span><span class="sxs-lookup"><span data-stu-id="fccbd-103">Count blank rows on sheets</span></span>

<span data-ttu-id="fccbd-104">Ce projet comprend deux scripts :</span><span class="sxs-lookup"><span data-stu-id="fccbd-104">This project includes two scripts:</span></span>

* <span data-ttu-id="fccbd-105">[Compter les lignes vides sur une feuille](#sample-code-count-blank-rows-on-a-given-sheet)donnée : parcourt la plage utilisée dans une feuille de calcul donnée et renvoie un nombre de lignes vide.</span><span class="sxs-lookup"><span data-stu-id="fccbd-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="fccbd-106">[Compter les lignes vides sur toutes les feuilles](#sample-code-count-blank-rows-on-all-sheets): parcourt la plage utilisée sur toutes les _feuilles_ de calcul et renvoie un nombre de lignes vide.</span><span class="sxs-lookup"><span data-stu-id="fccbd-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="fccbd-107">Pour notre script, une ligne vide est toute ligne sans données.</span><span class="sxs-lookup"><span data-stu-id="fccbd-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="fccbd-108">La ligne peut avoir une mise en forme.</span><span class="sxs-lookup"><span data-stu-id="fccbd-108">The row can have formatting.</span></span>

<span data-ttu-id="fccbd-109">_Cette feuille renvoie le nombre de 4 lignes vides_</span><span class="sxs-lookup"><span data-stu-id="fccbd-109">_This sheet returns count of 4 blank rows_</span></span>

![Données avec lignes vides](../../images/blank-rows.png)

<span data-ttu-id="fccbd-111">_Cette feuille renvoie le nombre de 0 lignes vides (toutes les lignes ont des données)_</span><span class="sxs-lookup"><span data-stu-id="fccbd-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

![Données sans lignes vides](../../images/no-blank-rows.png)

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="fccbd-113">Exemple de code : compter les lignes vides sur une feuille donnée</span><span class="sxs-lookup"><span data-stu-id="fccbd-113">Sample code: Count blank rows on a given sheet</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheet = workbook.getWorksheet('Sheet1'); 
  // Getting the active worksheet is not suitable for a script used by Power Automate.
  // const sheet = workbook.getActiveWorksheet();
  
  const range = sheet.getUsedRange(true); // Get value only.
  if (!range) {
    console.log(`No data on this sheet. `);
    return;
  }
  console.log(`Used range for the worksheet: ${range.getAddress()}`);
  const values = range.getValues();
  let emptyRows = 0;
  for (let row of values) {
    let len = 0; 
    for (let cell of row) {
      len = len + cell.toString().length;
    }
    if (len === 0) { 
      emptyRows++;
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="fccbd-114">Exemple de code : compter les lignes vides sur toutes les feuilles</span><span class="sxs-lookup"><span data-stu-id="fccbd-114">Sample code: Count blank rows on all sheets</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): number
{
  const sheets = workbook.getWorksheets();
  let emptyRows = 0;
  for (let sheet of sheets) { 
    const range = sheet.getUsedRange(true); // Get value only.
    if (!range) {
      console.log(`No data on this sheet. `);
      continue;
    }
    console.log(`Used range for the worksheet ${sheet.getName()}: ${range.getAddress()}`);
    const values = range.getValues();

    for (let row of values) {
      let len = 0;
      for (let cell of row) {
        len = len + cell.toString().length;
      }
      if (len === 0) {
        emptyRows++;
      }
    }
  }
  console.log(`Total empty row: ` + emptyRows);
  return emptyRows;
}
```

## <a name="use-with-power-automate"></a><span data-ttu-id="fccbd-115">Utilisation avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="fccbd-115">Use with Power Automate</span></span>

![Capture d’écran montrant comment configurer Power Automate](../../images/use-in-power-automate.png)
