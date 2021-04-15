---
title: Compter les lignes vides sur les feuilles
description: Découvrez comment utiliser les scripts Office pour détecter s'il existe des lignes vides au lieu de données dans les feuilles de calcul, puis signaler le nombre de lignes vierges à utiliser dans un flux Power Automate.
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: 088ab97c686484ca5c13c875b80431ac28d20736
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754830"
---
# <a name="count-blank-rows-on-sheets"></a><span data-ttu-id="fcfa0-103">Compter les lignes vides sur les feuilles</span><span class="sxs-lookup"><span data-stu-id="fcfa0-103">Count blank rows on sheets</span></span>

<span data-ttu-id="fcfa0-104">Ce projet comprend deux scripts :</span><span class="sxs-lookup"><span data-stu-id="fcfa0-104">This project includes two scripts:</span></span>

* <span data-ttu-id="fcfa0-105">[Compter les lignes vides sur une feuille](#sample-code-count-blank-rows-on-a-given-sheet)donnée : parcourt la plage utilisée dans une feuille de calcul donnée et renvoie un nombre de lignes vide.</span><span class="sxs-lookup"><span data-stu-id="fcfa0-105">[Count blank rows on a given sheet](#sample-code-count-blank-rows-on-a-given-sheet): Traverses the used range on a given worksheet and returns a blank row count.</span></span>
* <span data-ttu-id="fcfa0-106">[Compter les lignes vides sur toutes les feuilles](#sample-code-count-blank-rows-on-all-sheets): parcourt la plage utilisée sur toutes les _feuilles_ de calcul et renvoie un nombre de lignes vide.</span><span class="sxs-lookup"><span data-stu-id="fcfa0-106">[Count blank rows on all sheets](#sample-code-count-blank-rows-on-all-sheets): Traverses the used range on _all of the worksheets_ and returns a blank row count.</span></span>

> [!NOTE]
> <span data-ttu-id="fcfa0-107">Pour notre script, une ligne vide est toute ligne sans données.</span><span class="sxs-lookup"><span data-stu-id="fcfa0-107">For our script, a blank row is any row where there's no data.</span></span> <span data-ttu-id="fcfa0-108">La ligne peut avoir une mise en forme.</span><span class="sxs-lookup"><span data-stu-id="fcfa0-108">The row can have formatting.</span></span>

<span data-ttu-id="fcfa0-109">_Cette feuille renvoie le nombre de 4 lignes vides_</span><span class="sxs-lookup"><span data-stu-id="fcfa0-109">_This sheet returns count of 4 blank rows_</span></span>

:::image type="content" source="../../images/blank-rows.png" alt-text="Feuille de calcul affichant des données avec des lignes vides.":::

<span data-ttu-id="fcfa0-111">_Cette feuille renvoie le nombre de 0 lignes vides (toutes les lignes ont des données)_</span><span class="sxs-lookup"><span data-stu-id="fcfa0-111">_This sheet returns count of 0 blank rows (all rows have some data)_</span></span>

:::image type="content" source="../../images/no-blank-rows.png" alt-text="Feuille de calcul montrant les données sans lignes vides.":::

## <a name="sample-code-count-blank-rows-on-a-given-sheet"></a><span data-ttu-id="fcfa0-113">Exemple de code : compter les lignes vides sur une feuille donnée</span><span class="sxs-lookup"><span data-stu-id="fcfa0-113">Sample code: Count blank rows on a given sheet</span></span>

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

## <a name="sample-code-count-blank-rows-on-all-sheets"></a><span data-ttu-id="fcfa0-114">Exemple de code : compter les lignes vides sur toutes les feuilles</span><span class="sxs-lookup"><span data-stu-id="fcfa0-114">Sample code: Count blank rows on all sheets</span></span>

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

## <a name="use-with-power-automate"></a><span data-ttu-id="fcfa0-115">Utilisation avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="fcfa0-115">Use with Power Automate</span></span>

:::image type="content" source="../../images/use-in-power-automate.png" alt-text="Flux Power Automate montrant comment configurer pour exécuter un script Office.":::
