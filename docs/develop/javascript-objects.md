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
# <a name="use-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="60d7d-103">Utilisez des objets JavaScript intégrés dans les scripts Office’écriture</span><span class="sxs-lookup"><span data-stu-id="60d7d-103">Use built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="60d7d-104">JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, que vous scriptiez en JavaScript [ou TypeScript](../overview/code-editor-environment.md) (un superset de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="60d7d-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="60d7d-105">Cet article décrit comment vous pouvez utiliser certains des objets JavaScript intégrés dans les scripts Office pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="60d7d-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="60d7d-106">Pour une liste complète de tous les objets JavaScript intégrés, consultez l’article standard des objets [intégrés de Mozilla.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)</span><span class="sxs-lookup"><span data-stu-id="60d7d-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="60d7d-107">Tableau</span><span class="sxs-lookup"><span data-stu-id="60d7d-107">Array</span></span>

<span data-ttu-id="60d7d-108">[L’objet](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Array fournit une façon standardisée de travailler avec les tableaux de votre script.</span><span class="sxs-lookup"><span data-stu-id="60d7d-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="60d7d-109">Bien que les tableaux soient des constructions JavaScript standard, ils se rapportent Office scripts de deux manières majeures : les plages et les collections.</span><span class="sxs-lookup"><span data-stu-id="60d7d-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="work-with-ranges"></a><span data-ttu-id="60d7d-110">Travailler avec des plages</span><span class="sxs-lookup"><span data-stu-id="60d7d-110">Work with ranges</span></span>

<span data-ttu-id="60d7d-111">Les plages contiennent plusieurs tableaux bidimensionnels qui cartographient directement les cellules de cette plage.</span><span class="sxs-lookup"><span data-stu-id="60d7d-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="60d7d-112">Ces tableaux contiennent des informations spécifiques sur chaque cellule dans cette gamme.</span><span class="sxs-lookup"><span data-stu-id="60d7d-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="60d7d-113">Par exemple, `Range.getValues` renvoie toutes les valeurs de ces cellules (avec les lignes et les colonnes du tableau bidimensionnel cartographiant vers les lignes et les colonnes de cette sous-section de feuille de travail).</span><span class="sxs-lookup"><span data-stu-id="60d7d-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="60d7d-114">`Range.getFormulas` et `Range.getNumberFormats` sont d’autres méthodes fréquemment utilisées qui retournent des tableaux comme `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="60d7d-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="60d7d-115">Le script suivant recherche la **plage A1:D4 pour** n’importe quel format de numéro contenant un « $ ».</span><span class="sxs-lookup"><span data-stu-id="60d7d-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="60d7d-116">Le script définit la couleur de remplissage dans ces cellules à « jaune ».</span><span class="sxs-lookup"><span data-stu-id="60d7d-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="work-with-collections"></a><span data-ttu-id="60d7d-117">Travailler avec les collections</span><span class="sxs-lookup"><span data-stu-id="60d7d-117">Work with collections</span></span>

<span data-ttu-id="60d7d-118">De Excel objets sont contenus dans une collection.</span><span class="sxs-lookup"><span data-stu-id="60d7d-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="60d7d-119">La collection est gérée par l’Office Scripts et exposée sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="60d7d-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="60d7d-120">Par exemple, toutes [les](/javascript/api/office-scripts/excelscript/excelscript.shape) formes d’une feuille de travail sont contenues dans `Shape[]` une qui est retournée par la `Worksheet.getShapes` méthode.</span><span class="sxs-lookup"><span data-stu-id="60d7d-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="60d7d-121">Vous pouvez utiliser ce tableau pour lire les valeurs de la collection, ou vous pouvez accéder à des objets spécifiques à partir des méthodes de l’objet `get*` parent.</span><span class="sxs-lookup"><span data-stu-id="60d7d-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="60d7d-122">N’ajoutez pas ou ne supprimez pas manuellement les objets de ces tableaux de collection.</span><span class="sxs-lookup"><span data-stu-id="60d7d-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="60d7d-123">Utilisez les `add` méthodes sur les objets parent et les méthodes sur les objets de type `delete` collection.</span><span class="sxs-lookup"><span data-stu-id="60d7d-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="60d7d-124">Par exemple, ajoutez une table [à](/javascript/api/office-scripts/excelscript/excelscript.table) une feuille [de travail avec la](/javascript/api/office-scripts/excelscript/excelscript.worksheet) méthode et `Worksheet.addTable` supprimez `Table` l’utilisation `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="60d7d-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="60d7d-125">Le script suivant enregistre le type de chaque forme dans la feuille de travail actuelle.</span><span class="sxs-lookup"><span data-stu-id="60d7d-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="60d7d-126">Le script suivant supprime la forme la plus ancienne de la feuille de travail actuelle.</span><span class="sxs-lookup"><span data-stu-id="60d7d-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="60d7d-127">Date</span><span class="sxs-lookup"><span data-stu-id="60d7d-127">Date</span></span>

<span data-ttu-id="60d7d-128">[L’objet Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit une façon standardisée de travailler avec les dates de votre script.</span><span class="sxs-lookup"><span data-stu-id="60d7d-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="60d7d-129">`Date.now()` génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout de timetamps à la saisie de données de votre script.</span><span class="sxs-lookup"><span data-stu-id="60d7d-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="60d7d-130">Le script suivant ajoute la date actuelle à la feuille de travail.</span><span class="sxs-lookup"><span data-stu-id="60d7d-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="60d7d-131">Notez qu’en utilisant `toLocaleDateString` la méthode, Excel reconnaît la valeur comme une date et modifie automatiquement le format de nombre de la cellule.</span><span class="sxs-lookup"><span data-stu-id="60d7d-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="60d7d-132">La [section Travail avec dates](../resources/samples/excel-samples.md#dates) des échantillons contient plus de scripts liés à la date.</span><span class="sxs-lookup"><span data-stu-id="60d7d-132">The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="60d7d-133">Mathématiques</span><span class="sxs-lookup"><span data-stu-id="60d7d-133">Math</span></span>

<span data-ttu-id="60d7d-134">[L’objet](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) Mathématiques fournit des méthodes et des constantes pour les opérations mathématiques courantes.</span><span class="sxs-lookup"><span data-stu-id="60d7d-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="60d7d-135">Ceux-ci fournissent de nombreuses fonctions également disponibles Excel, sans avoir besoin d’utiliser le moteur de calcul du manuel.</span><span class="sxs-lookup"><span data-stu-id="60d7d-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="60d7d-136">Cela évite à votre script d’avoir à interroger le manuel, ce qui améliore les performances.</span><span class="sxs-lookup"><span data-stu-id="60d7d-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="60d7d-137">Le script suivant utilise `Math.min` pour trouver et enregistrer le plus petit nombre dans la gamme **A1:D4.**</span><span class="sxs-lookup"><span data-stu-id="60d7d-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="60d7d-138">Notez que cet échantillon suppose que toute la plage ne contient que des nombres, pas des chaînes.</span><span class="sxs-lookup"><span data-stu-id="60d7d-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="60d7d-139">L’utilisation de bibliothèques JavaScript externes n’est pas prise en charge</span><span class="sxs-lookup"><span data-stu-id="60d7d-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="60d7d-140">Office Les scripts ne supporte pas l’utilisation de bibliothèques externes tierces.</span><span class="sxs-lookup"><span data-stu-id="60d7d-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="60d7d-141">Votre script ne peut utiliser que les objets JavaScript intégrés et les Office API scripts.</span><span class="sxs-lookup"><span data-stu-id="60d7d-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="60d7d-142">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="60d7d-142">See also</span></span>

- [<span data-ttu-id="60d7d-143">Objets intégrés standard</span><span class="sxs-lookup"><span data-stu-id="60d7d-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="60d7d-144">Office Environnement scripts Code Editor</span><span class="sxs-lookup"><span data-stu-id="60d7d-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
