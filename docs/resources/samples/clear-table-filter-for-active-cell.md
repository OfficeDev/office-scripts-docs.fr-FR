---
title: Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives
description: Découvrez comment effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: f10e23b4ad948a28c5b749533ddedefe164d7142
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313889"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a><span data-ttu-id="d9a63-103">Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives</span><span class="sxs-lookup"><span data-stu-id="d9a63-103">Clear table column filter based on active cell location</span></span>

<span data-ttu-id="d9a63-104">Cet exemple permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement de la cellule active.</span><span class="sxs-lookup"><span data-stu-id="d9a63-104">This sample clears the table column filter based on the active cell location.</span></span> <span data-ttu-id="d9a63-105">Le script détecte si la cellule fait partie d’un tableau, détermine la colonne de tableau et clears any filter that are applied on it.</span><span class="sxs-lookup"><span data-stu-id="d9a63-105">The script detects if the cell is part of a table, determines the table column, and clears any filter that are applied on it.</span></span>

<span data-ttu-id="d9a63-106">Si vous souhaitez en savoir plus sur l’enregistrement du filtre avant de l’effacer (et appliquer à nouveau ultérieurement), voir Déplacer des lignes dans les tableaux en enregistreant des [filtres](move-rows-across-tables.md), un exemple plus avancé.</span><span class="sxs-lookup"><span data-stu-id="d9a63-106">If you wish to learn more about how to save the filter prior to clearing it (and re-apply later), see [Move rows across tables by saving filters](move-rows-across-tables.md), a more advanced sample.</span></span>

<span data-ttu-id="d9a63-107">_Avant d’effacer le filtre de colonne (notez la cellule active)_</span><span class="sxs-lookup"><span data-stu-id="d9a63-107">_Before clearing column filter (notice the active cell)_</span></span>

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Cellule active avant l’effacement du filtre de colonne.":::

<span data-ttu-id="d9a63-109">_Après l’effacement du filtre de colonne_</span><span class="sxs-lookup"><span data-stu-id="d9a63-109">_After clearing column filter_</span></span>

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Cellule active après l’effacement du filtre de colonne.":::

## <a name="sample-excel-file"></a><span data-ttu-id="d9a63-111">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="d9a63-111">Sample Excel file</span></span>

<span data-ttu-id="d9a63-112">Téléchargez <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> pour un livre de travail prêt à l’emploi.</span><span class="sxs-lookup"><span data-stu-id="d9a63-112">Download <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="d9a63-113">Ajoutez le script suivant pour essayer l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="d9a63-113">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a><span data-ttu-id="d9a63-114">Exemple de code : effacer le filtre de colonne de tableau en fonction de la cellule active</span><span class="sxs-lookup"><span data-stu-id="d9a63-114">Sample code: Clear table column filter based on active cell</span></span>

<span data-ttu-id="d9a63-115">Le script suivant permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives et peut être appliqué à Excel fichier avec une table.</span><span class="sxs-lookup"><span data-stu-id="d9a63-115">The following script clears the table column filter based on active cell location and can be applied to any Excel file with a table.</span></span>

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
