---
title: Utilisation d’objets JavaScript intégrés dans les scripts Office
description: Comment appeler des API JavaScript intégrées à partir d’un script Office dans Excel sur le Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545046"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Utiliser des objets JavaScript intégrés dans Office scripts

JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, que vous mentiez dans JavaScript ou [TypeScript](../overview/code-editor-environment.md) (un sur-ensemble de JavaScript). Cet article explique comment utiliser certains des objets JavaScript intégrés dans Office Scripts pour Excel sur le Web.

> [!NOTE]
> Pour obtenir la liste complète de tous les objets JavaScript intégrés, voir l’article sur les objets [intégrés Standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.

## <a name="array"></a>Tableau

[L’objet Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) fournit un moyen standard de travailler avec des tableaux dans votre script. Bien que les tableaux soient des constructions JavaScript standard, ils sont liés Office scripts de deux manières principales : les plages et les collections.

### <a name="work-with-ranges"></a>Travailler avec des plages

Les plages contiennent plusieurs tableaux à deux dimensions qui sont directement map faits sur les cellules de cette plage. Ces tableaux contiennent des informations spécifiques sur chaque cellule de cette plage. Par exemple, renvoie toutes les valeurs de ces cellules (avec les lignes et les colonnes du mappage de tableau à deux dimensions sur les lignes et les colonnes de cette sous-section de feuille `Range.getValues` de calcul). `Range.getFormulas` et `Range.getNumberFormats` sont d’autres méthodes fréquemment utilisées qui retournent des tableaux tels que `Range.getValues` .

Le script suivant recherche dans la plage **A1:D4** n’importe quel format de nombre contenant un « $ ». Le script définit la couleur de remplissage de ces cellules sur « jaune ».

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

### <a name="work-with-collections"></a>Travailler avec des collections

De Excel objets sont contenus dans une collection. La collection est gérée par l’API Office Scripts et exposée sous la mesure d’un tableau. Par exemple, toutes [les formes](/javascript/api/office-scripts/excelscript/excelscript.shape) d’une feuille de calcul sont contenues dans une forme `Shape[]` renvoyée par la `Worksheet.getShapes` méthode. Vous pouvez utiliser ce tableau pour lire les valeurs de la collection ou accéder à des objets spécifiques à partir des méthodes de l’objet `get*` parent.

> [!NOTE]
> N’ajoutez pas ou ne supprimez pas manuellement des objets de ces tableaux de collections. Utilisez les méthodes sur les objets parents et les méthodes sur les objets `add` `delete` de type collection. Par exemple, ajoutez un [tableau à](/javascript/api/office-scripts/excelscript/excelscript.table) une [feuille de](/javascript/api/office-scripts/excelscript/excelscript.worksheet) calcul avec la méthode `Worksheet.addTable` et supprimez l’utilisation. `Table` `Table.delete`

Le script suivant enregistre le type de chaque forme dans la feuille de calcul actuelle.

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

Le script suivant supprime la forme la plus ancienne dans la feuille de calcul actuelle.

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

[L’objet Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit un moyen standard de travailler avec les dates dans votre script. `Date.now()` génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout d’timestamps à l’entrée de données de votre script.

Le script suivant ajoute la date actuelle à la feuille de calcul. Notez qu’à l’aide de la méthode, Excel la valeur en tant que date et modifie automatiquement le format numérique `toLocaleDateString` de la cellule.

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

La section [Travailler avec les dates](../resources/samples/excel-samples.md#dates) des exemples contient davantage de scripts liés à la date.

## <a name="math"></a>Mathématiques

[L’objet Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fournit des méthodes et des constantes pour les opérations mathématiques courantes. Celles-ci fournissent de nombreuses fonctions également disponibles dans Excel, sans avoir à utiliser le moteur de calcul du workbook. Cela permet d’éviter à votre script d’avoir à interroger le workbook, ce qui améliore les performances.

Le script suivant utilise pour rechercher et enregistrer le plus petit nombre dans la plage `Math.min` **A1:D4.** Notez que cet exemple suppose que la plage entière contient uniquement des nombres, et non des chaînes.

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

Office Les scripts ne supportent pas l’utilisation de bibliothèques externes tierces. Votre script peut uniquement utiliser les objets JavaScript intégrés et les API Office Scripts.

## <a name="see-also"></a>Voir aussi

- [Objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Environnement d’éditeur de code scripts](../overview/code-editor-environment.md)
