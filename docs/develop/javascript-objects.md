---
title: Utilisation d’objets JavaScript intégrés dans les scripts Office
description: Comment appeler des API JavaScript intégrées à partir d’un script Office dans Excel sur le Web.
ms.date: 01/21/2020
localization_priority: Normal
ms.openlocfilehash: e0fcd98117125ead18e55675e195415ff59c0c5d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700201"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="f47e9-103">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="f47e9-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="f47e9-104">JavaScript fournit plusieurs objets intégrés que vous pouvez utiliser dans vos scripts Office, qu’il s’agisse de scripts JavaScript ou [dactylographiés](../overview/code-editor-environment.md) (un sur-ensemble de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="f47e9-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="f47e9-105">Cet article explique comment utiliser certains objets JavaScript intégrés dans les scripts Office pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="f47e9-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="f47e9-106">Pour obtenir la liste complète de tous les objets JavaScript intégrés, consultez l’article [objets intégrés standard](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) de Mozilla.</span><span class="sxs-lookup"><span data-stu-id="f47e9-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="f47e9-107">Tableau</span><span class="sxs-lookup"><span data-stu-id="f47e9-107">Array</span></span>

<span data-ttu-id="f47e9-108">L’objet [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) offre un moyen standardisé de travailler avec des tableaux dans votre script.</span><span class="sxs-lookup"><span data-stu-id="f47e9-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="f47e9-109">Bien que les tableaux soient des constructions JavaScript standard, ils sont liés aux scripts Office de deux manières principales : les plages et les collections.</span><span class="sxs-lookup"><span data-stu-id="f47e9-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="f47e9-110">Utilisation des plages</span><span class="sxs-lookup"><span data-stu-id="f47e9-110">Working with ranges</span></span>

<span data-ttu-id="f47e9-111">Les plages contiennent plusieurs tableaux à deux dimensions qui correspondent directement aux cellules de cette plage.</span><span class="sxs-lookup"><span data-stu-id="f47e9-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="f47e9-112">Elles incluent des propriétés telles `values`que `formulas`, et `numberFormat`.</span><span class="sxs-lookup"><span data-stu-id="f47e9-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="f47e9-113">Les propriétés de type tableau doivent être [chargées](scripting-fundamentals.md#sync-and-load) comme n’importe quelle autre propriété.</span><span class="sxs-lookup"><span data-stu-id="f47e9-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="f47e9-114">Le script suivant recherche dans la plage **a1 : D4** le format de nombre contenant un « $ ».</span><span class="sxs-lookup"><span data-stu-id="f47e9-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="f47e9-115">Le script définit la couleur de remplissage de ces cellules sur « jaune ».</span><span class="sxs-lookup"><span data-stu-id="f47e9-115">The script sets the fill color in those cells to "yellow".</span></span>

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

### <a name="working-with-collections"></a><span data-ttu-id="f47e9-116">Utilisation des collections</span><span class="sxs-lookup"><span data-stu-id="f47e9-116">Working with collections</span></span>

<span data-ttu-id="f47e9-117">De nombreux objets Excel sont contenus dans une collection.</span><span class="sxs-lookup"><span data-stu-id="f47e9-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="f47e9-118">Par exemple, toutes les [formes](/javascript/api/office-scripts/excel/excel.shape) d’une feuille de calcul sont contenues dans un `Worksheet.shapes` [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (en tant que propriété).</span><span class="sxs-lookup"><span data-stu-id="f47e9-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="f47e9-119">Chaque `*Collection` objet contient une `items` propriété qui est un tableau qui stocke les objets à l’intérieur de cette collection.</span><span class="sxs-lookup"><span data-stu-id="f47e9-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="f47e9-120">Cela peut être traité comme un tableau JavaScript normal, mais les éléments de la collection doivent d’abord être chargés.</span><span class="sxs-lookup"><span data-stu-id="f47e9-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="f47e9-121">Si vous devez utiliser une propriété sur chaque objet de la collection, utilisez une instruction de chargement hiérarchique (`items/propertyName`).</span><span class="sxs-lookup"><span data-stu-id="f47e9-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="f47e9-122">Le script suivant journalise le type de chaque forme dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="f47e9-122">The following script logs the type of every shape in the current worksheet.</span></span>

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

<span data-ttu-id="f47e9-123">Vous pouvez charger des objets individuels à partir d’une `getItem` collection `getItemAt` à l’aide des méthodes ou.</span><span class="sxs-lookup"><span data-stu-id="f47e9-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="f47e9-124">`getItem`Obtient un objet à l’aide d’un identificateur unique comme un nom (ces noms sont souvent spécifiés par votre script).</span><span class="sxs-lookup"><span data-stu-id="f47e9-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="f47e9-125">`getItemAt`Obtient un objet à l’aide de son index dans la collection.</span><span class="sxs-lookup"><span data-stu-id="f47e9-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="f47e9-126">L’appel doit être suivi d’une `await context.sync();` commande avant que l’objet puisse être utilisé.</span><span class="sxs-lookup"><span data-stu-id="f47e9-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="f47e9-127">Le script suivant supprime la forme la plus ancienne dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="f47e9-127">The following script deletes the oldest shape in the current worksheet.</span></span>

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

## <a name="date"></a><span data-ttu-id="f47e9-128">Date</span><span class="sxs-lookup"><span data-stu-id="f47e9-128">Date</span></span>

<span data-ttu-id="f47e9-129">L’objet [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fournit une méthode standardisée pour utiliser des dates dans votre script.</span><span class="sxs-lookup"><span data-stu-id="f47e9-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="f47e9-130">`Date.now()`génère un objet avec la date et l’heure actuelles, ce qui est utile lors de l’ajout d’horodatages à l’entrée de données de votre script.</span><span class="sxs-lookup"><span data-stu-id="f47e9-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="f47e9-131">Le script suivant ajoute la date actuelle à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="f47e9-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="f47e9-132">À l’aide de la `toLocaleDateString` méthode, Excel reconnaît la valeur comme une date et modifie automatiquement le format numérique de la cellule.</span><span class="sxs-lookup"><span data-stu-id="f47e9-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

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

## <a name="math"></a><span data-ttu-id="f47e9-133">Mathématiques</span><span class="sxs-lookup"><span data-stu-id="f47e9-133">Math</span></span>

<span data-ttu-id="f47e9-134">L’objet [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fournit des méthodes et des constantes pour les opérations mathématiques courantes.</span><span class="sxs-lookup"><span data-stu-id="f47e9-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="f47e9-135">Elles offrent de nombreuses fonctions également disponibles dans Excel, sans qu’il soit nécessaire d’utiliser le moteur de calcul du classeur.</span><span class="sxs-lookup"><span data-stu-id="f47e9-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="f47e9-136">Cela évite que votre script interroge le classeur, ce qui améliore les performances.</span><span class="sxs-lookup"><span data-stu-id="f47e9-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="f47e9-137">Le script suivant utilise `Math.min` pour rechercher et consigner le plus petit nombre de la plage **a1 : D4** .</span><span class="sxs-lookup"><span data-stu-id="f47e9-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="f47e9-138">Notez que cet exemple suppose que la plage entière ne contienne que des nombres, et non des chaînes.</span><span class="sxs-lookup"><span data-stu-id="f47e9-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="f47e9-139">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f47e9-139">See also</span></span>

- [<span data-ttu-id="f47e9-140">Objets intégrés standard</span><span class="sxs-lookup"><span data-stu-id="f47e9-140">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="f47e9-141">Environnement de l’éditeur de code des scripts Office</span><span class="sxs-lookup"><span data-stu-id="f47e9-141">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
