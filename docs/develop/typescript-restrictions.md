---
title: Restrictions TypeScript dans les scripts Office
description: Les spécificités du compilateur TypeScript et du linter utilisés par l’éditeur de code de scripts Office.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 41584ff23b333d17b2e267fdb3b0ec8741f3d203
ms.sourcegitcommit: df2b64603f91acb37bf95230efd538db0fbf9206
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/04/2021
ms.locfileid: "50099901"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restrictions TypeScript dans les scripts Office

Les scripts Office utilisent le langage TypeScript. Dans la plupart des cas, tout code TypeScript ou JavaScript fonctionne dans un script Office. Toutefois, il existe quelques restrictions appliquées par l’éditeur de code pour vous assurer que votre script fonctionne de manière cohérente et comme prévu avec votre workbook Excel.

## <a name="no-any-type-in-office-scripts"></a>Aucun type « any » dans les scripts Office

[L’écriture](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) de types est facultative dans TypeScript, car les types peuvent être déduits. Toutefois, Office Script requiert qu’une variable ne puisse pas être [de type n’importe quel](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Les scripts explicite et implicite `any` ne sont pas autorisés dans un script Office. Ces cas sont signalés comme des erreurs.

### <a name="explicit-any"></a>Explicite `any`

Vous ne pouvez pas déclarer explicitement une variable de type dans `any` Les scripts Office (c’est-à-dire, `let someVariable: any;` ). Le `any` type provoque des problèmes lors du traitement par Excel. Par exemple, il `Range` faut savoir qu’une valeur est `string` une , ou `number` `boolean` . Vous recevrez une erreur au moment de la compilation (une erreur avant l’exécution du script) si une variable est explicitement définie en tant que `any` type dans le script.

![Message explicite dans le texte de pointeur de l’éditeur de code](../images/explicit-any-editor-message.png)

![Erreur explicite dans la fenêtre de console](../images/explicit-any-error-message.png)

Dans la capture d’écran ci-dessus indique `[5, 16] Explicit Any is not allowed` que la ligne #5, la colonne #16 définit le `any` type. Cela vous permet de localiser l’erreur.

Pour contourner ce problème, définissez toujours le type de la variable. Si vous avez des doutes sur le type d’une variable, vous pouvez utiliser un [type d’union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Cela peut être utile pour les variables qui détiennent des valeurs, qui peuvent être de type , ou (le type des valeurs est une `Range` `string` union de `number` `boolean` `Range` celles-ci : `string | number | boolean` ).

### <a name="implicit-any"></a>Implicite `any`

Les types de variables TypeScript peuvent être [implicitement](( https://www.typescriptlang.org/docs/handbook/type-inference.html) définis. Si le compilateur TypeScript ne parvient pas à déterminer le type d’une variable (soit parce que le type n’est pas défini explicitement, soit parce que l’inférence de type n’est pas possible), il s’agit d’un implicite et vous recevrez une erreur au moment de la `any` compilation.

Le cas le plus courant sur tout implicite `any` se trouve dans une déclaration de variable, telle que `let value;` . Il existe deux façons d’éviter cela :

* Affectez la variable à un type implicitement identifiable ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).
* Tapez explicitement la variable ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Aucune classe ou interface Office Script n’hérite

Les classes et interfaces créées dans votre script Office ne peuvent pas étendre ou implémenter des classes ou interfaces [Office](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Scripts. En d’autres termes, rien dans l’espace de noms `ExcelScript` ne peut avoir de sous-classes ou de sous-polices.

## <a name="incompatible-typescript-functions"></a>Fonctions TypeScript incompatibles

Les API Office Scripts ne peuvent pas être utilisées dans les cas suivants :

* [Fonctions du générateur](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` n’est pas pris en charge

La fonction [d’eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript n’est pas prise en charge pour des raisons de sécurité.

## <a name="restricted-identifers"></a>Identifers restreints

Les mots suivants ne peuvent pas être utilisés comme identificateurs dans un script. Ce sont des termes réservés.

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a>Avertissements de performances

Le [linter](https://wikipedia.org/wiki/Lint_(software)) de l’éditeur de code avertit si le script peut avoir des problèmes de performances. Les cas et la façon de les contourner sont documentés dans [Améliorer les performances de vos scripts Office.](web-client-performance.md)

## <a name="external-api-calls"></a>Appels d’API externes

Pour plus [d’informations, voir](external-calls.md) la prise en charge des appels d’API externes dans Les scripts Office.

## <a name="see-also"></a>Voir aussi

* [Principes de base pour la rédaction de scripts Office en Excel sur le web](scripting-fundamentals.md)
* [Améliorer les performances de vos scripts Office](web-client-performance.md)
