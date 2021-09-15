---
title: Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives
description: Découvrez comment effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb01292017a027e41230d786337b5bf53293a20c
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332984"
---
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives

Cet exemple permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement de la cellule active. Le script détecte si la cellule fait partie d’un tableau, détermine la colonne de tableau et clears any filter that are applied on it.

Si vous souhaitez en savoir plus sur la façon d’enregistrer le filtre avant de l’effacer (et de le réapplager ultérieurement), voir Déplacer des lignes dans les tableaux en enregistreant des [filtres](move-rows-across-tables.md), un exemple plus avancé.

_Avant d’effacer le filtre de colonne (notez la cellule active)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Cellule active avant l’effacement du filtre de colonne.":::

_Après l’effacement du filtre de colonne_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Cellule active après l’effacement du filtre de colonne.":::

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez <a href="table-with-filter.xlsx">table-with-filter.xlsx</a> pour un livre de travail prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Exemple de code : effacer le filtre de colonne de tableau en fonction de la cellule active

Le script suivant permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives et peut être appliqué à Excel fichier avec un tableau.

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
