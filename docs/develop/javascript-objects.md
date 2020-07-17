---
title: Utilisation d’objets JavaScript intégrés dans les scripts Office
description: Comment appeler des API JavaScript intégrées à partir d’un script Office dans Excel sur le Web.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 1c8ac757574e8c4be64b373f8d4bf421ddfa0c79
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999259"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="18786-103">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="18786-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="18786-104">JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, qu’il s’agisse de scripts JavaScript ou [dactylographiés](../overview/code-editor-environment.md) (un sur-ensemble de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="18786-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="18786-105">Cet article explique comment utiliser certains objets JavaScript intégrés dans les scripts Office pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="18786-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="18786-106">Pour obtenir la liste complète de tous les objets JavaScript intégrés, consultez l’article [objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.</span><span class="sxs-lookup"><span data-stu-id="18786-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="18786-107">Tableau</span><span class="sxs-lookup"><span data-stu-id="18786-107">Array</span></span>

<span data-ttu-id="18786-108">L’objet [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) offre un moyen standardisé de travailler avec des tableaux dans votre script.</span><span class="sxs-lookup"><span data-stu-id="18786-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="18786-109">Bien que les tableaux soient des constructions JavaScript standard, ils sont liés aux scripts Office de deux manières principales : les plages et les collections.</span><span class="sxs-lookup"><span data-stu-id="18786-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="18786-110">Utilisation des plages</span><span class="sxs-lookup"><span data-stu-id="18786-110">Working with ranges</span></span>

<span data-ttu-id="18786-111">Les plages contiennent plusieurs tableaux à deux dimensions qui correspondent directement aux cellules de cette plage.</span><span class="sxs-lookup"><span data-stu-id="18786-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="18786-112">Ces tableaux contiennent des informations spécifiques sur chaque cellule de cette plage.</span><span class="sxs-lookup"><span data-stu-id="18786-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="18786-113">Par exemple, `Range.getValues` renvoie toutes les valeurs de ces cellules (avec les lignes et les colonnes du mappage de tableau à deux dimensions sur les lignes et les colonnes de cette sous-section de feuille de calcul).</span><span class="sxs-lookup"><span data-stu-id="18786-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="18786-114">`Range.getFormulas`et `Range.getNumberFormats` sont d’autres méthodes fréquemment utilisées qui retournent des tableaux comme `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="18786-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="18786-115">Le script suivant recherche dans la plage **a1 : D4** le format de nombre contenant un « $ ».</span><span class="sxs-lookup"><span data-stu-id="18786-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="18786-116">Le script définit la couleur de remplissage de ces cellules sur « jaune ».</span><span class="sxs-lookup"><span data-stu-id="18786-116">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="18786-117">Utilisation des collections</span><span class="sxs-lookup"><span data-stu-id="18786-117">Working with collections</span></span>

<span data-ttu-id="18786-118">De nombreux objets Excel sont contenus dans une collection.</span><span class="sxs-lookup"><span data-stu-id="18786-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="18786-119">La collection est gérée par l’API de scripts Office et exposée sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="18786-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="18786-120">Par exemple, toutes les [formes](/javascript/api/office-scripts/excelscript/excelscript.shape) d’une feuille de calcul sont contenues dans un `Shape[]` qui est renvoyé par la `Worksheet.getShapes` méthode.</span><span class="sxs-lookup"><span data-stu-id="18786-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="18786-121">Vous pouvez utiliser ce tableau pour lire des valeurs à partir de la collection, ou pour accéder à des objets spécifiques à partir des méthodes de l’objet parent `get*` .</span><span class="sxs-lookup"><span data-stu-id="18786-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="18786-122">N’ajoutez pas ou ne supprimez pas manuellement des objets de ces tableaux de collections.</span><span class="sxs-lookup"><span data-stu-id="18786-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="18786-123">Utilisez les `add` méthodes sur les objets parents et les `delete` méthodes sur les objets de type collection.</span><span class="sxs-lookup"><span data-stu-id="18786-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="18786-124">Par exemple, ajoutez une [table](/javascript/api/office-scripts/excelscript/excelscript.table) à une [feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) avec la `Worksheet.addTable` méthode et supprimez l' `Table` using `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="18786-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="18786-125">Le script suivant journalise le type de chaque forme dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="18786-125">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="18786-126">Le script suivant supprime la forme la plus ancienne dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="18786-126">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="18786-127">Date</span><span class="sxs-lookup"><span data-stu-id="18786-127">Date</span></span>

<span data-ttu-id="18786-128">L’objet [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit une méthode standardisée pour utiliser des dates dans votre script.</span><span class="sxs-lookup"><span data-stu-id="18786-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="18786-129">`Date.now()`génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout d’horodatages à l’entrée de données de votre script.</span><span class="sxs-lookup"><span data-stu-id="18786-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="18786-130">Le script suivant ajoute la date actuelle à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="18786-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="18786-131">À l’aide de la `toLocaleDateString` méthode, Excel reconnaît la valeur comme une date et modifie automatiquement le format numérique de la cellule.</span><span class="sxs-lookup"><span data-stu-id="18786-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

<span data-ttu-id="18786-132">La section [utiliser les dates](../resources/excel-samples.md#work-with-dates) des exemples contient davantage de scripts liés à la date.</span><span class="sxs-lookup"><span data-stu-id="18786-132">The [Work with dates](../resources/excel-samples.md#work-with-dates) section of the samples has more Date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="18786-133">Mathématiques</span><span class="sxs-lookup"><span data-stu-id="18786-133">Math</span></span>

<span data-ttu-id="18786-134">L’objet [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fournit des méthodes et des constantes pour les opérations mathématiques courantes.</span><span class="sxs-lookup"><span data-stu-id="18786-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="18786-135">Elles offrent de nombreuses fonctions également disponibles dans Excel, sans qu’il soit nécessaire d’utiliser le moteur de calcul du classeur.</span><span class="sxs-lookup"><span data-stu-id="18786-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="18786-136">Cela évite que votre script interroge le classeur, ce qui améliore les performances.</span><span class="sxs-lookup"><span data-stu-id="18786-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="18786-137">Le script suivant utilise `Math.min` pour rechercher et consigner le plus petit nombre de la plage **a1 : D4** .</span><span class="sxs-lookup"><span data-stu-id="18786-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="18786-138">Notez que cet exemple suppose que la plage entière ne contienne que des nombres, et non des chaînes.</span><span class="sxs-lookup"><span data-stu-id="18786-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="18786-139">L’utilisation de bibliothèques JavaScript externes n’est pas prise en charge</span><span class="sxs-lookup"><span data-stu-id="18786-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="18786-140">Les scripts Office ne prennent pas en charge l’utilisation de bibliothèques tierces externes.</span><span class="sxs-lookup"><span data-stu-id="18786-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="18786-141">Votre script peut uniquement utiliser les objets JavaScript intégrés et les API de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="18786-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="18786-142">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="18786-142">See also</span></span>

- [<span data-ttu-id="18786-143">Objets intégrés standard</span><span class="sxs-lookup"><span data-stu-id="18786-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="18786-144">Environnement de l’éditeur de code des scripts Office</span><span class="sxs-lookup"><span data-stu-id="18786-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
