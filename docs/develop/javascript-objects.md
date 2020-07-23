---
title: Utilisation d’objets JavaScript intégrés dans les scripts Office
description: Comment appeler des API JavaScript intégrées à partir d’un script Office dans Excel sur le Web.
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: 4bb5fb5444887005ececbbfdf0130cba3784e0c4
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229595"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Utilisation d’objets JavaScript intégrés dans les scripts Office

JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, qu’il s’agisse de scripts JavaScript ou [dactylographiés](../overview/code-editor-environment.md) (un sur-ensemble de JavaScript). Cet article explique comment utiliser certains objets JavaScript intégrés dans les scripts Office pour Excel sur le Web.

> [!NOTE]
> Pour obtenir la liste complète de tous les objets JavaScript intégrés, consultez l’article [objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.

## <a name="array"></a>Tableau

L’objet [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) offre un moyen standardisé de travailler avec des tableaux dans votre script. Bien que les tableaux soient des constructions JavaScript standard, ils sont liés aux scripts Office de deux manières principales : les plages et les collections.

### <a name="working-with-ranges"></a>Utilisation des plages

Les plages contiennent plusieurs tableaux à deux dimensions qui correspondent directement aux cellules de cette plage. Ces tableaux contiennent des informations spécifiques sur chaque cellule de cette plage. Par exemple, `Range.getValues` renvoie toutes les valeurs de ces cellules (avec les lignes et les colonnes du mappage de tableau à deux dimensions sur les lignes et les colonnes de cette sous-section de feuille de calcul). `Range.getFormulas`et `Range.getNumberFormats` sont d’autres méthodes fréquemment utilisées qui retournent des tableaux comme `Range.getValues` .

Le script suivant recherche dans la plage **a1 : D4** le format de nombre contenant un « $ ». Le script définit la couleur de remplissage de ces cellules sur « jaune ».

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="working-with-collections"></a>Utilisation des collections

De nombreux objets Excel sont contenus dans une collection. La collection est gérée par l’API de scripts Office et exposée sous forme de tableau. Par exemple, toutes les [formes](/javascript/api/office-scripts/excelscript/excelscript.shape) d’une feuille de calcul sont contenues dans un `Shape[]` qui est renvoyé par la `Worksheet.getShapes` méthode. Vous pouvez utiliser ce tableau pour lire des valeurs à partir de la collection, ou pour accéder à des objets spécifiques à partir des méthodes de l’objet parent `get*` .

> [!NOTE]
> N’ajoutez pas ou ne supprimez pas manuellement des objets de ces tableaux de collections. Utilisez les `add` méthodes sur les objets parents et les `delete` méthodes sur les objets de type collection. Par exemple, ajoutez une [table](/javascript/api/office-scripts/excelscript/excelscript.table) à une [feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) avec la `Worksheet.addTable` méthode et supprimez l' `Table` using `Table.delete` .

Le script suivant journalise le type de chaque forme dans la feuille de calcul active.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

Le script suivant supprime la forme la plus ancienne dans la feuille de calcul active.

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a>Date

L’objet [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit une méthode standardisée pour utiliser des dates dans votre script. `Date.now()`génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout d’horodatages à l’entrée de données de votre script.

Le script suivant ajoute la date actuelle à la feuille de calcul. À l’aide de la `toLocaleDateString` méthode, Excel reconnaît la valeur comme une date et modifie automatiquement le format numérique de la cellule.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

La section [utiliser les dates](../resources/excel-samples.md#dates) des exemples contient davantage de scripts liés à la date.

## <a name="math"></a>Mathématiques

L’objet [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fournit des méthodes et des constantes pour les opérations mathématiques courantes. Elles offrent de nombreuses fonctions également disponibles dans Excel, sans qu’il soit nécessaire d’utiliser le moteur de calcul du classeur. Cela évite que votre script interroge le classeur, ce qui améliore les performances.

Le script suivant utilise `Math.min` pour rechercher et consigner le plus petit nombre de la plage **a1 : D4** . Notez que cet exemple suppose que la plage entière ne contienne que des nombres, et non des chaînes.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>L’utilisation de bibliothèques JavaScript externes n’est pas prise en charge

Les scripts Office ne prennent pas en charge l’utilisation de bibliothèques tierces externes. Votre script peut uniquement utiliser les objets JavaScript intégrés et les API de scripts Office.

## <a name="see-also"></a>Voir aussi

- [Objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Environnement de l’éditeur de code des scripts Office](../overview/code-editor-environment.md)
