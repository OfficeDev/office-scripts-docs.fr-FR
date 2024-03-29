---
title: Supprimer les filtres de la colonne du tableau
description: Découvrez comment effacer le filtre de colonne de table en fonction de l’emplacement de cellule actif.
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: e016f7f2af9e7553229f3b3b19007e011879de8e
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572520"
---
# <a name="remove-table-column-filters"></a>Supprimer les filtres de la colonne du tableau

Cet exemple supprime les filtres d’une colonne de table, en fonction de l’emplacement de la cellule active. Le script détecte si la cellule fait partie d’une table, détermine la colonne de table et efface tous les filtres qui y sont appliqués.

Si vous souhaitez en savoir plus sur l’enregistrement du filtre avant de l’effacer (et le réappliquer ultérieurement), consultez [Déplacer des lignes entre les tables en enregistrant les filtres](move-rows-across-tables.md), un exemple plus avancé.

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez [table-with-filter.xlsx](table-with-filter.xlsx) pour un classeur prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-clear-table-column-filter-based-on-active-cell"></a>Exemple de code : Effacer le filtre de colonne de table en fonction de la cellule active

Le script suivant efface le filtre de colonne de table en fonction de l’emplacement de cellule actif et peut être appliqué à n’importe quel fichier Excel avec une table.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  const cell = workbook.getActiveCell();

  // Get the tables associated with that cell.
  // Since tables can't overlap, this will be one table at most.
  const currentTable = cell.getTables()[0];

  // If there is no table on the selection, end the script.
  if (!currentTable) {
    console.log("The selection is not in a table.");
    return;
  }

  // Get the table header above the current cell by referencing its column.
  const entireColumn = cell.getEntireColumn();
  const intersect = entireColumn.getIntersection(currentTable.getRange());
  const headerCellValue = intersect.getCell(0, 0).getValue() as string;

  // Get the TableColumn object matching that header.
  const tableColumn = currentTable.getColumnByName(headerCellValue);

  // Clear the filters on that table column.
  tableColumn.getFilter().clear();
}
```

## <a name="before-clearing-column-filter-notice-the-active-cell"></a>Avant d’effacer le filtre de colonne (notez la cellule active)

:::image type="content" source="../../images/before-filter-applied.png" alt-text="Cellule active avant d’effacer le filtre de colonne.":::

## <a name="after-clearing-column-filter"></a>Après l’effacement du filtre de colonne

:::image type="content" source="../../images/after-filter-cleared.png" alt-text="Cellule active après l’effacement du filtre de colonne.":::
