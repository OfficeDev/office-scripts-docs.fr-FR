---
title: Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives
description: Découvrez comment effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 06fba191a79f4641d4d1017bda332c7559b50e6d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285926"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="5bdee-103">Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives</span><span class="sxs-lookup"><span data-stu-id="5bdee-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="5bdee-104">Cet exemple permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement de la cellule active.</span><span class="sxs-lookup"><span data-stu-id="5bdee-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="5bdee-105">Le script détecte si la cellule fait partie d’un tableau, détermine la colonne de tableau et clears any filter that are applied on it.</span><span class="sxs-lookup"><span data-stu-id="5bdee-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="5bdee-106">Si vous souhaitez en savoir plus sur la façon d’enregistrer le filtre avant de l’effacer (et de le réapplager ultérieurement), voir Déplacer des lignes dans les tableaux en enregistreant des [filtres](move-rows-across-tables.md), un exemple plus avancé.</span><span class="sxs-lookup"><span data-stu-id="5bdee-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="5bdee-107">_Avant d’effacer le filtre de colonne (notez la cellule active)_</span><span class="sxs-lookup"><span data-stu-id="5bdee-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Cellule active avant l’effacement du filtre de colonne":::

<span data-ttu-id="5bdee-109">_Après l’effacement du filtre de colonne_</span><span class="sxs-lookup"><span data-stu-id="5bdee-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Cellule active après l’effacement du filtre de colonne":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="5bdee-111">Exemple de code : effacer le filtre de colonne de tableau en fonction de la cellule active</span><span class="sxs-lookup"><span data-stu-id="5bdee-111">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="5bdee-112">Le script suivant permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives et peut être appliqué à Excel fichier avec un tableau.</span><span class="sxs-lookup"><span data-stu-id="5bdee-112">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span> <span data-ttu-id="5bdee-113">Pour plus de commodité, vous pouvez télécharger et utiliser <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span><span class="sxs-lookup"><span data-stu-id="5bdee-113">For convenience, you can download and use <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active cell.
    const cell = workbook.getActiveCell();

    // Get all tables associated with that cell.
    const tables = cell.getTables();
    
    // If there is no table on the selection, end the script.
    if (tables.length !== 1) {
      console.log("The selection is not in a table.");
      return;
    }

    // Get the first table associated with the active cell.
    const currentTable = tables[0];

    // Log key information about the table.
    console.log(currentTable.getName());
    console.log(currentTable.getRange().getAddress());

    // Get the table header above the current cell by referencing its column.
    const entireColumn = cell.getEntireColumn();
    const intersect = entireColumn.getIntersection(currentTable.getRange());
    console.log(intersect.getAddress());

    const headerCellValue = intersect.getCell(0,0).getValue() as string;
    console.log(headerCellValue);

    // Get the TableColumn object matching that header.
    const tableColumn = currentTable.getColumnByName(headerCellValue);

    // Clear the filter on that table column.
    tableColumn.getFilter().clear();
}
```
