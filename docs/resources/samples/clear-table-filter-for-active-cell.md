---
title: Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives
description: Découvrez comment effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives.
ms.date: 03/04/2021
localization_priority: Normal
ms.openlocfilehash: bbca4adce1de2cfade2c4f84273bf0bc06b5cc4b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232500"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="3b03e-103">Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives</span><span class="sxs-lookup"><span data-stu-id="3b03e-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="3b03e-104">Cet exemple permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement de la cellule active.</span><span class="sxs-lookup"><span data-stu-id="3b03e-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="3b03e-105">Le script détecte si la cellule fait partie d’un tableau, détermine la colonne de tableau et clears any filter that are applied on it.</span><span class="sxs-lookup"><span data-stu-id="3b03e-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="3b03e-106">Si vous souhaitez en savoir plus sur la façon d’enregistrer le filtre avant de l’effacer (et de le réapplager ultérieurement), voir Déplacer des lignes dans les tableaux en enregistreant des [filtres](move-rows-across-tables.md), un exemple plus avancé.</span><span class="sxs-lookup"><span data-stu-id="3b03e-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="3b03e-107">_Avant d’effacer le filtre de colonne (notez la cellule active)_</span><span class="sxs-lookup"><span data-stu-id="3b03e-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Cellule active avant l’effacement du filtre de colonne":::

<span data-ttu-id="3b03e-109">_Après l’effacement du filtre de colonne_</span><span class="sxs-lookup"><span data-stu-id="3b03e-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Cellule active après l’effacement du filtre de colonne":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="3b03e-111">Exemple de code : effacer le filtre de colonne de tableau en fonction de la cellule active</span><span class="sxs-lookup"><span data-stu-id="3b03e-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="3b03e-112">Le script suivant permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives et peut être appliqué à Excel fichier avec un tableau.</span><span class="sxs-lookup"><span data-stu-id="3b03e-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="3b03e-113">Pour plus de commodité, vous pouvez télécharger et utiliser <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span><span class="sxs-lookup"><span data-stu-id="3b03e-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, return/exit.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get table (since it is already determined that there is only
    // a single table part of the selection).
    const currentTable = tables[0];

    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    const entireCol = cell.getEntireColumn();
    const intersect = entireCol.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get column.
    const col = currentTable.getColumnByName(headerCellValue);

    // Clear filter.
    col.getFilter().clear();
}
```
