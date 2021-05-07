---
title: Utilisation d’objets JavaScript intégrés dans les scripts Office
description: Comment appeler des API JavaScript intégrées à partir d’un script Office dans Excel sur le Web.
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: e3b36265f235678eee18fbf369058b165da46210
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232402"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="68ee6-103">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="68ee6-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="68ee6-104">JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, que vous mentiez dans JavaScript ou [TypeScript](../overview/code-editor-environment.md) (un sur-ensemble de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="68ee6-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="68ee6-105">Cet article explique comment utiliser certains des objets JavaScript intégrés dans Office Scripts pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="68ee6-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="68ee6-106">Pour obtenir la liste complète de tous les objets JavaScript intégrés, voir l’article sur les objets [intégrés Standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.</span><span class="sxs-lookup"><span data-stu-id="68ee6-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="68ee6-107">Tableau</span><span class="sxs-lookup"><span data-stu-id="68ee6-107">Array</span></span>

<span data-ttu-id="68ee6-108">[L’objet Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) fournit un moyen standard de travailler avec des tableaux dans votre script.</span><span class="sxs-lookup"><span data-stu-id="68ee6-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="68ee6-109">Bien que les tableaux soient des constructions JavaScript standard, ils sont liés Office scripts de deux manières principales : plages et collections.</span><span class="sxs-lookup"><span data-stu-id="68ee6-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="68ee6-110">Travailler avec des plages</span><span class="sxs-lookup"><span data-stu-id="68ee6-110">Working with ranges</span></span>

<span data-ttu-id="68ee6-111">Les plages contiennent plusieurs tableaux à deux dimensions qui sont directement map faits sur les cellules de cette plage.</span><span class="sxs-lookup"><span data-stu-id="68ee6-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="68ee6-112">Ces tableaux contiennent des informations spécifiques sur chaque cellule de cette plage.</span><span class="sxs-lookup"><span data-stu-id="68ee6-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="68ee6-113">Par exemple, renvoie toutes les valeurs de ces cellules (avec les lignes et les colonnes du mappage de tableau à deux dimensions sur les lignes et les colonnes de cette sous-section de feuille `Range.getValues` de calcul).</span><span class="sxs-lookup"><span data-stu-id="68ee6-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="68ee6-114">`Range.getFormulas` et `Range.getNumberFormats` sont d’autres méthodes fréquemment utilisées qui retournent des tableaux tels que `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="68ee6-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="68ee6-115">Le script suivant recherche dans la plage **A1:D4** n’importe quel format de nombre contenant un « $ ».</span><span class="sxs-lookup"><span data-stu-id="68ee6-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="68ee6-116">Le script définit la couleur de remplissage de ces cellules sur « jaune ».</span><span class="sxs-lookup"><span data-stu-id="68ee6-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="68ee6-117">Travailler avec des collections</span><span class="sxs-lookup"><span data-stu-id="68ee6-117">Working with collections</span></span>

<span data-ttu-id="68ee6-118">De Excel objets sont contenus dans une collection.</span><span class="sxs-lookup"><span data-stu-id="68ee6-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="68ee6-119">La collection est gérée par l’API Office Scripts et exposée sous la mesure d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="68ee6-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="68ee6-120">Par exemple, toutes [les formes](/javascript/api/office-scripts/excelscript/excelscript.shape) d’une feuille de calcul sont contenues dans une forme `Shape[]` renvoyée par la `Worksheet.getShapes` méthode.</span><span class="sxs-lookup"><span data-stu-id="68ee6-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="68ee6-121">Vous pouvez utiliser ce tableau pour lire les valeurs de la collection ou accéder à des objets spécifiques à partir des méthodes de l’objet `get*` parent.</span><span class="sxs-lookup"><span data-stu-id="68ee6-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="68ee6-122">N’ajoutez pas ou ne supprimez pas manuellement des objets de ces tableaux de collections.</span><span class="sxs-lookup"><span data-stu-id="68ee6-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="68ee6-123">Utilisez les méthodes sur les objets parents et les méthodes sur les objets `add` `delete` de type collection.</span><span class="sxs-lookup"><span data-stu-id="68ee6-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="68ee6-124">Par exemple, ajoutez un [tableau à](/javascript/api/office-scripts/excelscript/excelscript.table) une [feuille de](/javascript/api/office-scripts/excelscript/excelscript.worksheet) calcul avec la méthode `Worksheet.addTable` et supprimez l’utilisation. `Table` `Table.delete`</span><span class="sxs-lookup"><span data-stu-id="68ee6-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="68ee6-125">Le script suivant enregistre le type de chaque forme dans la feuille de calcul actuelle.</span><span class="sxs-lookup"><span data-stu-id="68ee6-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="68ee6-126">Le script suivant supprime la forme la plus ancienne dans la feuille de calcul actuelle.</span><span class="sxs-lookup"><span data-stu-id="68ee6-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="68ee6-127">Date</span><span class="sxs-lookup"><span data-stu-id="68ee6-127">Date</span></span>

<span data-ttu-id="68ee6-128">[L’objet Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit un moyen standard de travailler avec les dates dans votre script.</span><span class="sxs-lookup"><span data-stu-id="68ee6-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="68ee6-129">`Date.now()` génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout d’timestamps à l’entrée de données de votre script.</span><span class="sxs-lookup"><span data-stu-id="68ee6-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="68ee6-130">Le script suivant ajoute la date actuelle à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="68ee6-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="68ee6-131">Notez qu’à l’aide de la méthode, Excel la valeur en tant que date et modifie automatiquement le format numérique `toLocaleDateString` de la cellule.</span><span class="sxs-lookup"><span data-stu-id="68ee6-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="68ee6-132">La section [Travailler avec les dates](../resources/samples/excel-samples.md#dates) des exemples contient davantage de scripts liés à la date.</span><span class="sxs-lookup"><span data-stu-id="68ee6-132">The [Work with dates](../resources/samples/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="68ee6-133">Mathématiques</span><span class="sxs-lookup"><span data-stu-id="68ee6-133">Math</span></span>

<span data-ttu-id="68ee6-134">[L’objet Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fournit des méthodes et des constantes pour les opérations mathématiques courantes.</span><span class="sxs-lookup"><span data-stu-id="68ee6-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="68ee6-135">Celles-ci fournissent de nombreuses fonctions également disponibles dans Excel, sans avoir besoin d’utiliser le moteur de calcul dubook.</span><span class="sxs-lookup"><span data-stu-id="68ee6-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="68ee6-136">Cela permet d’éviter à votre script d’avoir à interroger le workbook, ce qui améliore les performances.</span><span class="sxs-lookup"><span data-stu-id="68ee6-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="68ee6-137">Le script suivant utilise pour rechercher et enregistrer le plus petit nombre dans la plage `Math.min` **A1:D4.**</span><span class="sxs-lookup"><span data-stu-id="68ee6-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="68ee6-138">Notez que cet exemple suppose que la plage entière contient uniquement des nombres, et non des chaînes.</span><span class="sxs-lookup"><span data-stu-id="68ee6-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="68ee6-139">L’utilisation de bibliothèques JavaScript externes n’est pas prise en charge</span><span class="sxs-lookup"><span data-stu-id="68ee6-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="68ee6-140">Office Les scripts ne supportent pas l’utilisation de bibliothèques externes tierces.</span><span class="sxs-lookup"><span data-stu-id="68ee6-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="68ee6-141">Votre script peut uniquement utiliser les objets JavaScript intégrés et les API Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="68ee6-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="68ee6-142">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="68ee6-142">See also</span></span>

- [<span data-ttu-id="68ee6-143">Objets intégrés standard</span><span class="sxs-lookup"><span data-stu-id="68ee6-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="68ee6-144">Office Environnement d’éditeur de code scripts</span><span class="sxs-lookup"><span data-stu-id="68ee6-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
