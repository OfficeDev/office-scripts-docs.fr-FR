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
# <a name="typescript-restrictions-in-office-scripts"></a>Restrictions typescript dans Office scripts

Office Les scripts utilisent le langage TypeScript. Pour la plupart, n’importe quel code TypeScript ou JavaScript fonctionnera dans Office scripts. Toutefois, il existe quelques restrictions appliquées par l’éditeur de code pour s’assurer que votre script fonctionne de façon cohérente et comme prévu avec votre Excel de travail.

## <a name="no-any-type-in-office-scripts"></a>Pas de type « n’importe quel » Office scripts

Les [types d’écriture](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) sont facultatifs dans TypeScript, car les types peuvent être déduits. Toutefois, Office scripts exige qu’une variable ne peut pas être [de type tout](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Les scripts explicites `any` et implicites ne sont pas autorisés Office scripts. Ces cas sont signalés comme des erreurs.

### <a name="explicit-any"></a>explicite `any`

Vous ne pouvez pas déclarer explicitement une variable de type dans les `any` scripts Office (c’est-à-dire). `let someVariable: any;` Le `any` type provoque des problèmes lorsqu’il est traité Excel. Par exemple, un `Range` besoin de savoir qu’une valeur est un `string` , ou `number` `boolean` . Vous recevrez une erreur de temps de compilation (une erreur avant d’exécuter le script) si une variable est explicitement définie comme `any` le type dans le script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Le message explicite « n’importe quel » dans le texte stationnaire de l’éditeur de code":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="L’erreur explicite « n’importe quel » dans la fenêtre de la console":::

Dans la capture `[5, 16] Explicit Any is not allowed` d’écran précédente indique que #5 ligne, colonne #16 définit `any` type. Cela vous aide à localiser l’erreur.

Pour contourner ce problème, définissez toujours le type de variable. Si vous n’êtes pas certain du type de variable, vous pouvez utiliser un [type d’union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Cela peut être utile pour les variables qui détiennent `Range` des valeurs, qui peuvent être de type `string` , ou `number` `boolean` (le type de valeurs est une union `Range` de ceux qui: `string | number | boolean` ).

### <a name="implicit-any"></a>implicite `any`

Type Les types variables descript peuvent [être définis implicitement.](https://www.typescriptlang.org/docs/handbook/type-inference.html) Si le compilateur TypeScript n’est pas en mesure de déterminer le type d’une variable (soit parce que le type n’est pas défini explicitement, soit que l’inférence de type n’est pas possible), alors il s’agit d’une erreur implicite et vous `any` recevrez une erreur de temps de compilation.

Le cas le plus courant sur tout implicite `any` est dans une déclaration variable, telle que `let value;` . Il y a deux façons d’éviter cela :

* Attribuez la variable à un type implicitement identifiable ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).
* Tapez explicitement la variable ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Pas d’Office classes ou interfaces script

Les classes et interfaces créées dans votre script Office peuvent pas étendre [ou implémenter](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) des classes Office scripts ou des interfaces. En d’autres termes, rien dans `ExcelScript` l’espace nominatif ne peut avoir de sous-classes ou de sous-interfaces.

## <a name="incompatible-typescript-functions"></a>Fonctions incompatibles TypeScript

Office Les API scripts ne peuvent pas être utilisées dans les éléments suivants :

* [Fonctions du générateur](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` n’est pas pris en charge

La fonction eval JavaScript [n’est](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) pas prise en charge pour des raisons de sécurité.

## <a name="restricted-identifers"></a>Identifères restreints

Les mots suivants ne peuvent pas être utilisés comme identificateurs dans un script. Ce sont des conditions réservées.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Seules les fonctions fléchées dans les rappels de tableau

Vos scripts ne peuvent utiliser des fonctions [fléchées que](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) lorsqu’ils fournissent des arguments de rappel [pour les méthodes Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Vous ne pouvez pas transmettre une sorte d’identificateur ou de fonction « traditionnelle » à ces méthodes.

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

## <a name="performance-warnings"></a>Avertissements de performance

L’inter de [l’éditeur de](https://wikipedia.org/wiki/Lint_(software)) code donne des avertissements si le script peut avoir des problèmes de performances. Les cas et la façon de travailler autour d’eux sont [documentés dans Améliorer les performances de vos Office Scripts](web-client-performance.md).

## <a name="external-api-calls"></a>Appels API externes

Consultez [le support d’appel api externe dans Office scripts pour](external-calls.md) plus d’informations.

## <a name="see-also"></a>Voir aussi

* [Principes de base pour la rédaction de scripts Office en Excel sur le web](scripting-fundamentals.md)
* [Améliorez les performances de vos scripts Office’argent](web-client-performance.md)
