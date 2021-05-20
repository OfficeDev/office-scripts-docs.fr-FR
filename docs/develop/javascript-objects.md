---
title: Utilisation d’objets JavaScript intégrés dans les scripts Office
description: Comment appeler les API JavaScript intégrées à partir d’un script Office dans Excel sur le Web.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 680dd326e357bd06e2fc66cba5bd6745bbd33c24
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545046"
---
# <a name="use-built-in-javascript-objects-in-office-scripts"></a>Utilisez des objets JavaScript intégrés dans les scripts Office’écriture

JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, que vous scriptiez en JavaScript [ou TypeScript](../overview/code-editor-environment.md) (un superset de JavaScript). Cet article décrit comment vous pouvez utiliser certains des objets JavaScript intégrés dans les scripts Office pour Excel sur le Web.

> [!NOTE]
> Pour une liste complète de tous les objets JavaScript intégrés, consultez l’article standard des objets [intégrés de Mozilla.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)

## <a name="array"></a>Tableau

[L’objet](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Array fournit une façon standardisée de travailler avec les tableaux de votre script. Bien que les tableaux soient des constructions JavaScript standard, ils se rapportent Office scripts de deux manières majeures : les plages et les collections.

### <a name="work-with-ranges"></a>Travailler avec des plages

Les plages contiennent plusieurs tableaux bidimensionnels qui cartographient directement les cellules de cette plage. Ces tableaux contiennent des informations spécifiques sur chaque cellule dans cette gamme. Par exemple, `Range.getValues` renvoie toutes les valeurs de ces cellules (avec les lignes et les colonnes du tableau bidimensionnel cartographiant vers les lignes et les colonnes de cette sous-section de feuille de travail). `Range.getFormulas` et `Range.getNumberFormats` sont d’autres méthodes fréquemment utilisées qui retournent des tableaux comme `Range.getValues` .

Le script suivant recherche la **plage A1:D4 pour** n’importe quel format de numéro contenant un « $ ». Le script définit la couleur de remplissage dans ces cellules à « jaune ».

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

### <a name="work-with-collections"></a>Travailler avec les collections

De Excel objets sont contenus dans une collection. La collection est gérée par l’Office Scripts et exposée sous forme de tableau. Par exemple, toutes [les](/javascript/api/office-scripts/excelscript/excelscript.shape) formes d’une feuille de travail sont contenues dans `Shape[]` une qui est retournée par la `Worksheet.getShapes` méthode. Vous pouvez utiliser ce tableau pour lire les valeurs de la collection, ou vous pouvez accéder à des objets spécifiques à partir des méthodes de l’objet `get*` parent.

> [!NOTE]
> N’ajoutez pas ou ne supprimez pas manuellement les objets de ces tableaux de collection. Utilisez les `add` méthodes sur les objets parent et les méthodes sur les objets de type `delete` collection. Par exemple, ajoutez une table [à](/javascript/api/office-scripts/excelscript/excelscript.table) une feuille [de travail avec la](/javascript/api/office-scripts/excelscript/excelscript.worksheet) méthode et `Worksheet.addTable` supprimez `Table` l’utilisation `Table.delete` .

Le script suivant enregistre le type de chaque forme dans la feuille de travail actuelle.

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

Le script suivant supprime la forme la plus ancienne de la feuille de travail actuelle.

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

[L’objet Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit une façon standardisée de travailler avec les dates de votre script. `Date.now()` génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout de timetamps à la saisie de données de votre script.

Le script suivant ajoute la date actuelle à la feuille de travail. Notez qu’en utilisant `toLocaleDateString` la méthode, Excel reconnaît la valeur comme une date et modifie automatiquement le format de nombre de la cellule.

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

La [section Travail avec dates](../resources/samples/excel-samples.md#dates) des échantillons contient plus de scripts liés à la date.

## <a name="math"></a>Mathématiques

[L’objet](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) Mathématiques fournit des méthodes et des constantes pour les opérations mathématiques courantes. Ceux-ci fournissent de nombreuses fonctions également disponibles Excel, sans avoir besoin d’utiliser le moteur de calcul du manuel. Cela évite à votre script d’avoir à interroger le manuel, ce qui améliore les performances.

Le script suivant utilise `Math.min` pour trouver et enregistrer le plus petit nombre dans la gamme **A1:D4.** Notez que cet échantillon suppose que toute la plage ne contient que des nombres, pas des chaînes.

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

Office Les scripts ne supporte pas l’utilisation de bibliothèques externes tierces. Votre script ne peut utiliser que les objets JavaScript intégrés et les Office API scripts.

## <a name="see-also"></a>Voir aussi

- [Objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Office Environnement scripts Code Editor](../overview/code-editor-environment.md)
