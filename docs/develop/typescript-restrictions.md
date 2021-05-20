---
title: Restrictions typescript dans Office scripts
description: Les spécificités du compilateur typeScript et linter utilisé par l’éditeur Office Code scripts.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: a4198e0e56224ac5da89e89c43c8d2f3ef44d6d7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545018"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="f6044-103">Restrictions typescript dans Office scripts</span><span class="sxs-lookup"><span data-stu-id="f6044-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="f6044-104">Office Les scripts utilisent le langage TypeScript.</span><span class="sxs-lookup"><span data-stu-id="f6044-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="f6044-105">Pour la plupart, n’importe quel code TypeScript ou JavaScript fonctionnera dans Office scripts.</span><span class="sxs-lookup"><span data-stu-id="f6044-105">For the most part, any TypeScript or JavaScript code will work in Office Scripts.</span></span> <span data-ttu-id="f6044-106">Toutefois, il existe quelques restrictions appliquées par l’éditeur de code pour s’assurer que votre script fonctionne de façon cohérente et comme prévu avec votre Excel de travail.</span><span class="sxs-lookup"><span data-stu-id="f6044-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="f6044-107">Pas de type « n’importe quel » Office scripts</span><span class="sxs-lookup"><span data-stu-id="f6044-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="f6044-108">Les [types d’écriture](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sont facultatifs dans TypeScript, car les types peuvent être déduits.</span><span class="sxs-lookup"><span data-stu-id="f6044-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="f6044-109">Toutefois, Office scripts exige qu’une variable ne peut pas être [de type tout](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span><span class="sxs-lookup"><span data-stu-id="f6044-109">However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="f6044-110">Les scripts explicites `any` et implicites ne sont pas autorisés Office scripts.</span><span class="sxs-lookup"><span data-stu-id="f6044-110">Both explicit and implicit `any` are not allowed in Office Scripts.</span></span> <span data-ttu-id="f6044-111">Ces cas sont signalés comme des erreurs.</span><span class="sxs-lookup"><span data-stu-id="f6044-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="f6044-112">explicite `any`</span><span class="sxs-lookup"><span data-stu-id="f6044-112">Explicit `any`</span></span>

<span data-ttu-id="f6044-113">Vous ne pouvez pas déclarer explicitement une variable de type dans les `any` scripts Office (c’est-à-dire). `let someVariable: any;`</span><span class="sxs-lookup"><span data-stu-id="f6044-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="f6044-114">Le `any` type provoque des problèmes lorsqu’il est traité Excel.</span><span class="sxs-lookup"><span data-stu-id="f6044-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="f6044-115">Par exemple, un `Range` besoin de savoir qu’une valeur est un `string` , ou `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="f6044-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="f6044-116">Vous recevrez une erreur de temps de compilation (une erreur avant d’exécuter le script) si une variable est explicitement définie comme `any` le type dans le script.</span><span class="sxs-lookup"><span data-stu-id="f6044-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Le message explicite « n’importe quel » dans le texte stationnaire de l’éditeur de code":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="L’erreur explicite « n’importe quel » dans la fenêtre de la console":::

<span data-ttu-id="f6044-119">Dans la capture `[5, 16] Explicit Any is not allowed` d’écran précédente indique que #5 ligne, colonne #16 définit `any` type.</span><span class="sxs-lookup"><span data-stu-id="f6044-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="f6044-120">Cela vous aide à localiser l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f6044-120">This helps you locate the error.</span></span>

<span data-ttu-id="f6044-121">Pour contourner ce problème, définissez toujours le type de variable.</span><span class="sxs-lookup"><span data-stu-id="f6044-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="f6044-122">Si vous n’êtes pas certain du type de variable, vous pouvez utiliser un [type d’union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="f6044-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="f6044-123">Cela peut être utile pour les variables qui détiennent `Range` des valeurs, qui peuvent être de type `string` , ou `number` `boolean` (le type de valeurs est une union `Range` de ceux qui: `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="f6044-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="f6044-124">implicite `any`</span><span class="sxs-lookup"><span data-stu-id="f6044-124">Implicit `any`</span></span>

<span data-ttu-id="f6044-125">Type Les types variables descript peuvent [être définis implicitement.](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="f6044-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="f6044-126">Si le compilateur TypeScript n’est pas en mesure de déterminer le type d’une variable (soit parce que le type n’est pas défini explicitement, soit que l’inférence de type n’est pas possible), alors il s’agit d’une erreur implicite et vous `any` recevrez une erreur de temps de compilation.</span><span class="sxs-lookup"><span data-stu-id="f6044-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="f6044-127">Le cas le plus courant sur tout implicite `any` est dans une déclaration variable, telle que `let value;` .</span><span class="sxs-lookup"><span data-stu-id="f6044-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="f6044-128">Il y a deux façons d’éviter cela :</span><span class="sxs-lookup"><span data-stu-id="f6044-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="f6044-129">Attribuez la variable à un type implicitement identifiable ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="f6044-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="f6044-130">Tapez explicitement la variable ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="f6044-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="f6044-131">Pas d’Office classes ou interfaces script</span><span class="sxs-lookup"><span data-stu-id="f6044-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="f6044-132">Les classes et interfaces créées dans votre script Office peuvent pas étendre [ou implémenter](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) des classes Office scripts ou des interfaces.</span><span class="sxs-lookup"><span data-stu-id="f6044-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="f6044-133">En d’autres termes, rien dans `ExcelScript` l’espace nominatif ne peut avoir de sous-classes ou de sous-interfaces.</span><span class="sxs-lookup"><span data-stu-id="f6044-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="f6044-134">Fonctions incompatibles TypeScript</span><span class="sxs-lookup"><span data-stu-id="f6044-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="f6044-135">Office Les API scripts ne peuvent pas être utilisées dans les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="f6044-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="f6044-136">Fonctions du générateur</span><span class="sxs-lookup"><span data-stu-id="f6044-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="f6044-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="f6044-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="f6044-138">`eval` n’est pas pris en charge</span><span class="sxs-lookup"><span data-stu-id="f6044-138">`eval` is not supported</span></span>

<span data-ttu-id="f6044-139">La fonction eval JavaScript [n’est](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) pas prise en charge pour des raisons de sécurité.</span><span class="sxs-lookup"><span data-stu-id="f6044-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="f6044-140">Identifères restreints</span><span class="sxs-lookup"><span data-stu-id="f6044-140">Restricted identifers</span></span>

<span data-ttu-id="f6044-141">Les mots suivants ne peuvent pas être utilisés comme identificateurs dans un script.</span><span class="sxs-lookup"><span data-stu-id="f6044-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="f6044-142">Ce sont des conditions réservées.</span><span class="sxs-lookup"><span data-stu-id="f6044-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="f6044-143">Seules les fonctions fléchées dans les rappels de tableau</span><span class="sxs-lookup"><span data-stu-id="f6044-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="f6044-144">Vos scripts ne peuvent utiliser des fonctions [fléchées que](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) lorsqu’ils fournissent des arguments de rappel [pour les méthodes Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array)</span><span class="sxs-lookup"><span data-stu-id="f6044-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="f6044-145">Vous ne pouvez pas transmettre une sorte d’identificateur ou de fonction « traditionnelle » à ces méthodes.</span><span class="sxs-lookup"><span data-stu-id="f6044-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## <a name="performance-warnings"></a><span data-ttu-id="f6044-146">Avertissements de performance</span><span class="sxs-lookup"><span data-stu-id="f6044-146">Performance warnings</span></span>

<span data-ttu-id="f6044-147">L’inter de [l’éditeur de](https://wikipedia.org/wiki/Lint_(software)) code donne des avertissements si le script peut avoir des problèmes de performances.</span><span class="sxs-lookup"><span data-stu-id="f6044-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="f6044-148">Les cas et la façon de travailler autour d’eux sont [documentés dans Améliorer les performances de vos Office Scripts](web-client-performance.md).</span><span class="sxs-lookup"><span data-stu-id="f6044-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="f6044-149">Appels API externes</span><span class="sxs-lookup"><span data-stu-id="f6044-149">External API calls</span></span>

<span data-ttu-id="f6044-150">Consultez [le support d’appel api externe dans Office scripts pour](external-calls.md) plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="f6044-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="f6044-151">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f6044-151">See also</span></span>

* [<span data-ttu-id="f6044-152">Principes de base pour la rédaction de scripts Office en Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="f6044-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="f6044-153">Améliorez les performances de vos scripts Office’argent</span><span class="sxs-lookup"><span data-stu-id="f6044-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
