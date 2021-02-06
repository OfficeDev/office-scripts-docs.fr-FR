---
title: Restrictions TypeScript dans les scripts Office
description: Les spécificités du compilateur TypeScript et du linter utilisés par l’éditeur de code de scripts Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: d67e208561ce6ddd706d4c80cf29d2f013a32032
ms.sourcegitcommit: 98c7bc26f51dc8427669c571135c503d73bcee4c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/06/2021
ms.locfileid: "50125933"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="552b7-103">Restrictions TypeScript dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="552b7-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="552b7-104">Les scripts Office utilisent le langage TypeScript.</span><span class="sxs-lookup"><span data-stu-id="552b7-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="552b7-105">Dans la plupart des cas, tout code TypeScript ou JavaScript fonctionne dans un script Office.</span><span class="sxs-lookup"><span data-stu-id="552b7-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="552b7-106">Toutefois, il existe quelques restrictions appliquées par l’éditeur de code pour vous assurer que votre script fonctionne de manière cohérente et comme prévu avec votre workbook Excel.</span><span class="sxs-lookup"><span data-stu-id="552b7-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="552b7-107">Aucun type « any » dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="552b7-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="552b7-108">[L’écriture](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) de types est facultative dans TypeScript, car les types peuvent être déduits.</span><span class="sxs-lookup"><span data-stu-id="552b7-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="552b7-109">Toutefois, Office Script requiert qu’une variable ne puisse pas être [de type n’importe quel](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span><span class="sxs-lookup"><span data-stu-id="552b7-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="552b7-110">Les scripts explicites `any` et implicites ne sont pas autorisés dans un script Office.</span><span class="sxs-lookup"><span data-stu-id="552b7-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="552b7-111">Ces cas sont signalés comme des erreurs.</span><span class="sxs-lookup"><span data-stu-id="552b7-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="552b7-112">Explicite `any`</span><span class="sxs-lookup"><span data-stu-id="552b7-112">Explicit `any`</span></span>

<span data-ttu-id="552b7-113">Vous ne pouvez pas déclarer explicitement une variable de type dans `any` Les scripts Office (c’est-à-dire, `let someVariable: any;` ).</span><span class="sxs-lookup"><span data-stu-id="552b7-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="552b7-114">Le `any` type provoque des problèmes lors du traitement par Excel.</span><span class="sxs-lookup"><span data-stu-id="552b7-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="552b7-115">Par exemple, il `Range` faut savoir qu’une valeur est `string` un , ou `number` `boolean` .</span><span class="sxs-lookup"><span data-stu-id="552b7-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="552b7-116">Vous recevrez une erreur de compilation (une erreur avant l’exécution du script) si une variable est explicitement définie en tant que `any` type dans le script.</span><span class="sxs-lookup"><span data-stu-id="552b7-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![Message explicite dans le texte de pointeur de l’éditeur de code](../images/explicit-any-editor-message.png)

![Erreur explicite dans la fenêtre de console](../images/explicit-any-error-message.png)

<span data-ttu-id="552b7-119">Dans la capture d’écran ci-dessus indique `[5, 16] Explicit Any is not allowed` que la ligne #5, la colonne #16 définit le `any` type.</span><span class="sxs-lookup"><span data-stu-id="552b7-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="552b7-120">Cela vous permet de localiser l’erreur.</span><span class="sxs-lookup"><span data-stu-id="552b7-120">This helps you locate the error.</span></span>

<span data-ttu-id="552b7-121">Pour contourner ce problème, définissez toujours le type de la variable.</span><span class="sxs-lookup"><span data-stu-id="552b7-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="552b7-122">Si vous avez des doutes sur le type d’une variable, vous pouvez utiliser un [type d’union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)</span><span class="sxs-lookup"><span data-stu-id="552b7-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="552b7-123">Cela peut être utile pour les variables qui détiennent des valeurs, qui peuvent être de type , ou (le type des valeurs est une `Range` `string` union de `number` `boolean` `Range` celles-ci : `string | number | boolean` ).</span><span class="sxs-lookup"><span data-stu-id="552b7-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="552b7-124">Implicite `any`</span><span class="sxs-lookup"><span data-stu-id="552b7-124">Implicit `any`</span></span>

<span data-ttu-id="552b7-125">Les types de variables TypeScript peuvent être [implicitement définis.](https://www.typescriptlang.org/docs/handbook/type-inference.html)</span><span class="sxs-lookup"><span data-stu-id="552b7-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="552b7-126">Si le compilateur TypeScript ne parvient pas à déterminer le type d’une variable (soit parce que le type n’est pas défini explicitement, soit parce que l’inférence de type n’est pas possible), il s’agit d’un implicite et vous recevrez une erreur au moment de la `any` compilation.</span><span class="sxs-lookup"><span data-stu-id="552b7-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="552b7-127">Le cas le plus courant sur tout implicite `any` se trouve dans une déclaration de variable, telle que `let value;` .</span><span class="sxs-lookup"><span data-stu-id="552b7-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="552b7-128">Il existe deux façons d’éviter cela :</span><span class="sxs-lookup"><span data-stu-id="552b7-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="552b7-129">Affectez la variable à un type implicitement identifiable ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).</span><span class="sxs-lookup"><span data-stu-id="552b7-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="552b7-130">Tapez explicitement la variable ( `let value: number;` )</span><span class="sxs-lookup"><span data-stu-id="552b7-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="552b7-131">Aucune classe ou interface Office Script n’hérite</span><span class="sxs-lookup"><span data-stu-id="552b7-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="552b7-132">Les classes et interfaces créées dans votre script Office ne peuvent pas étendre ou implémenter des classes ou interfaces [Office](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Scripts.</span><span class="sxs-lookup"><span data-stu-id="552b7-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="552b7-133">En d’autres termes, rien dans l’espace `ExcelScript` de noms ne peut avoir de sous-classes ou de sous-polices.</span><span class="sxs-lookup"><span data-stu-id="552b7-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="552b7-134">Fonctions TypeScript incompatibles</span><span class="sxs-lookup"><span data-stu-id="552b7-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="552b7-135">Les API Office Scripts ne peuvent pas être utilisées dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="552b7-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="552b7-136">Fonctions du générateur</span><span class="sxs-lookup"><span data-stu-id="552b7-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="552b7-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="552b7-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="552b7-138">`eval` n’est pas pris en charge</span><span class="sxs-lookup"><span data-stu-id="552b7-138">`eval` is not supported</span></span>

<span data-ttu-id="552b7-139">La fonction [d’eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript n’est pas prise en charge pour des raisons de sécurité.</span><span class="sxs-lookup"><span data-stu-id="552b7-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="552b7-140">Identifers restreints</span><span class="sxs-lookup"><span data-stu-id="552b7-140">Restricted identifers</span></span>

<span data-ttu-id="552b7-141">Les mots suivants ne peuvent pas être utilisés comme identificateurs dans un script.</span><span class="sxs-lookup"><span data-stu-id="552b7-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="552b7-142">Ce sont des termes réservés.</span><span class="sxs-lookup"><span data-stu-id="552b7-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="552b7-143">Avertissements de performances</span><span class="sxs-lookup"><span data-stu-id="552b7-143">Performance warnings</span></span>

<span data-ttu-id="552b7-144">Le [linter](https://wikipedia.org/wiki/Lint_(software)) de l’éditeur de code avertit si le script peut avoir des problèmes de performances.</span><span class="sxs-lookup"><span data-stu-id="552b7-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="552b7-145">Les cas et la façon de les contourner sont documentés dans [Améliorer les performances de vos scripts Office.](web-client-performance.md)</span><span class="sxs-lookup"><span data-stu-id="552b7-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="552b7-146">Appels d’API externes</span><span class="sxs-lookup"><span data-stu-id="552b7-146">External API calls</span></span>

<span data-ttu-id="552b7-147">Pour plus [d’informations, voir](external-calls.md) la prise en charge des appels d’API externes dans Les scripts Office.</span><span class="sxs-lookup"><span data-stu-id="552b7-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="552b7-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="552b7-148">See also</span></span>

* [<span data-ttu-id="552b7-149">Principes de base pour la rédaction de scripts Office en Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="552b7-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="552b7-150">Améliorer les performances de vos scripts Office</span><span class="sxs-lookup"><span data-stu-id="552b7-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)