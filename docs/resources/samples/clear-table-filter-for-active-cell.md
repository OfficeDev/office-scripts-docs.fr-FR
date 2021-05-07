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
# <a name="clear-table-column-filter-based-on-active-cell-location"></a>Effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives

Cet exemple permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement de la cellule active. Le script détecte si la cellule fait partie d’un tableau, détermine la colonne de tableau et clears any filter that are applied on it.

Si vous souhaitez en savoir plus sur la façon d’enregistrer le filtre avant de l’effacer (et de le réapplager ultérieurement), voir Déplacer des lignes dans les tableaux en enregistreant des [filtres](move-rows-across-tables.md), un exemple plus avancé.

_Avant d’effacer le filtre de colonne (notez la cellule active)_

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Cellule active avant l’effacement du filtre de colonne":::

_Après l’effacement du filtre de colonne_

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Cellule active après l’effacement du filtre de colonne":::

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Exemple de code : effacer le filtre de colonne de tableau en fonction de la cellule active

Le script suivant permet d’effacer le filtre de colonne de tableau en fonction de l’emplacement des cellules actives et peut être appliqué à Excel fichier avec un tableau. Pour plus de commodité, vous pouvez télécharger et utiliser <a href="table-with-filter.xlsx">table-with-filter.xlsx</a>.

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
