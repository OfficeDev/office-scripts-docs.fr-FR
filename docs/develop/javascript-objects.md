---
title: Utilisation d’objets JavaScript intégrés dans les scripts Office
description: Comment appeler des API JavaScript intégrées à partir d’un script Office dans Excel sur le Web.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: a4b698215edea5f266e159fee0e08690904dd379
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191012"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Utilisation d’objets JavaScript intégrés dans les scripts Office

JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, qu’il s’agisse de scripts JavaScript ou [dactylographiés](../overview/code-editor-environment.md) (un sur-ensemble de JavaScript). Cet article explique comment utiliser certains objets JavaScript intégrés dans les scripts Office pour Excel sur le Web.

> [!NOTE]
> Pour obtenir la liste complète de tous les objets JavaScript intégrés, consultez l’article [objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.

## <a name="array"></a>Tableau

L’objet [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) offre un moyen standardisé de travailler avec des tableaux dans votre script. Bien que les tableaux soient des constructions JavaScript standard, ils sont liés aux scripts Office de deux manières principales : les plages et les collections.

### <a name="working-with-ranges"></a>Utilisation des plages

Les plages contiennent plusieurs tableaux à deux dimensions qui correspondent directement aux cellules de cette plage. Elles incluent des propriétés telles `values`que `formulas`, et `numberFormat`. Les propriétés de type tableau doivent être [chargées](scripting-fundamentals.md#sync-and-load) comme n’importe quelle autre propriété.

Le script suivant recherche dans la plage **a1 : D4** le format de nombre contenant un « $ ». Le script définit la couleur de remplissage de ces cellules sur « jaune ».

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Utilisation des collections

De nombreux objets Excel sont contenus dans une collection. Par exemple, toutes les [formes](/javascript/api/office-scripts/excel/excel.shape) d’une feuille de calcul sont contenues dans un `Worksheet.shapes` [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (en tant que propriété). Chaque `*Collection` objet contient une `items` propriété qui est un tableau qui stocke les objets à l’intérieur de cette collection. Cela peut être traité comme un tableau JavaScript normal, mais les éléments de la collection doivent d’abord être chargés. Si vous devez utiliser une propriété sur chaque objet de la collection, utilisez une instruction de chargement hiérarchique (`items/propertyName`).

Le script suivant journalise le type de chaque forme dans la feuille de calcul active.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

Vous pouvez charger des objets individuels à partir d’une `getItem` collection `getItemAt` à l’aide des méthodes ou. `getItem`Obtient un objet à l’aide d’un identificateur unique comme un nom (ces noms sont souvent spécifiés par votre script). `getItemAt`Obtient un objet à l’aide de son index dans la collection. L’appel doit être suivi d’une `await context.sync();` commande avant que l’objet puisse être utilisé.

Le script suivant supprime la forme la plus ancienne dans la feuille de calcul active.

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Date

L’objet [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit une méthode standardisée pour utiliser des dates dans votre script. `Date.now()`génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout d’horodatages à l’entrée de données de votre script.

Le script suivant ajoute la date actuelle à la feuille de calcul. À l’aide de la `toLocaleDateString` méthode, Excel reconnaît la valeur comme une date et modifie automatiquement le format numérique de la cellule.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

La section [utiliser les dates](../resources/excel-samples.md#work-with-dates) des exemples contient davantage de scripts liés à la date.

## <a name="math"></a>Mathématiques

L’objet [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fournit des méthodes et des constantes pour les opérations mathématiques courantes. Elles offrent de nombreuses fonctions également disponibles dans Excel, sans qu’il soit nécessaire d’utiliser le moteur de calcul du classeur. Cela évite que votre script interroge le classeur, ce qui améliore les performances.

Le script suivant utilise `Math.min` pour rechercher et consigner le plus petit nombre de la plage **a1 : D4** . Notez que cet exemple suppose que la plage entière ne contienne que des nombres, et non des chaînes.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```

## <a name="see-also"></a>Voir aussi

- [Objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Environnement de l’éditeur de code des scripts Office](../overview/code-editor-environment.md)
