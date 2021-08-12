---
title: Restrictions TypeScript dans Office scripts
description: Les spécificités du compilateur TypeScript et du linter utilisés par l’éditeur de code Office Scripts.
ms.date: 07/14/2021
localization_priority: Normal
ms.openlocfilehash: ea7b9e34b09409fbe7b4cfdab221a59d50246773167fbe6d1c64bbd61fd0b2df
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847039"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Restrictions TypeScript dans Office scripts

Office Les scripts utilisent le langage TypeScript. Dans la plupart des cas, tout code TypeScript ou JavaScript fonctionne dans Office scripts. Toutefois, il existe quelques restrictions appliquées par l’éditeur de code pour vous assurer que votre script fonctionne de manière cohérente et conformément à vos Excel de travail.

## <a name="no-any-type-in-office-scripts"></a>Aucun type « any » dans Office Scripts

[L’écriture](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) de types est facultative dans TypeScript, car les types peuvent être déduits. Toutefois, Office Scripts requiert qu’une variable ne puisse pas être [de type n’importe quel](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Les scripts explicites et implicites ne sont pas `any` autorisés Office scripts. Ces cas sont signalés comme des erreurs.

### <a name="explicit-any"></a>Explicite `any`

Vous ne pouvez pas déclarer explicitement une variable de type Office `any` scripts (c’est-à-dire, `let value: any;` ). Le `any` type provoque des problèmes lorsqu’il est Excel. Par exemple, il `Range` faut savoir qu’une valeur est `string` une , ou `number` `boolean` . Vous recevrez une erreur au moment de la compilation (une erreur avant l’exécution du script) si une variable est explicitement définie en tant que `any` type dans le script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Message explicite « any » dans le texte de pointeur de l’éditeur de code.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Erreur explicite « any » dans la fenêtre de console.":::

Dans la capture d’écran précédente, indique que la ligne `[2, 14] Explicit Any is not allowed` #2, la colonne #14 définit le `any` type. Cela vous permet de localiser l’erreur.

Pour contourner ce problème, définissez toujours le type de la variable. Si vous avez des doutes sur le type d’une variable, vous pouvez utiliser un [type d’union.](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html) Cela peut être utile pour les variables qui détiennent des valeurs, qui peuvent être de type , ou (le type des valeurs est une `Range` `string` union de `number` `boolean` `Range` celles-ci : `string | number | boolean` ).

### <a name="implicit-any"></a>Implicite `any`

Les types de variables TypeScript peuvent être [implicitement définis.](https://www.typescriptlang.org/docs/handbook/type-inference.html) Si le compilateur TypeScript ne parvient pas à déterminer le type d’une variable (soit parce que le type n’est pas défini explicitement, soit parce que l’inférence de type n’est pas possible), il s’agit d’un implicite et vous recevrez une erreur au moment de la `any` compilation.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Message implicite « any » dans le texte de pointeur de l’éditeur de code.":::

Le cas le plus courant sur tout implicite `any` se trouve dans une déclaration de variable, telle que `let value;` . Il existe deux façons d’éviter cela :

* Affectez la variable à un type implicitement identifiable ( `let value = 5;` ou `let value = workbook.getWorksheet();` ).
* Tapez explicitement la variable ( `let value: number;` )

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>Pas d’héritage Office classes ou interfaces de script

Les classes et interfaces créées dans votre [](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) script Office ne peuvent pas étendre ou implémenter Office classes ou interfaces scripts. En d’autres termes, rien dans l’espace `ExcelScript` de noms ne peut avoir de sous-classes ou de sous-polices.

## <a name="incompatible-typescript-functions"></a>Fonctions TypeScript incompatibles

Office Les API scripts ne peuvent pas être utilisées dans les cas suivants :

* [Fonctions du générateur](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` n’est pas pris en charge

La fonction [d’val](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) JavaScript n’est pas prise en charge pour des raisons de sécurité.

## <a name="restricted-identifers"></a>Identifers restreints

Les mots suivants ne peuvent pas être utilisés comme identificateurs dans un script. Ce sont des termes réservés.

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>Fonctions de flèche uniquement dans les rappels de tableau

Vos scripts peuvent uniquement utiliser des [fonctions de direction lors](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) de la fourniture d’arguments de rappel pour les méthodes [Array.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) Vous ne pouvez pas transmettre un type d’identificateur ou de fonction « traditionnelle » à ces méthodes.

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

## <a name="unions-of-excelscript-types-and-user-defined-types-arent-supported"></a>Les `ExcelScript` syndicats de types et les types définis par l’utilisateur ne sont pas pris en charge

Office Les scripts sont convertis au moment de l’runtime de blocs de code synchrones en blocs de code asynchrone. La communication avec le workbook par le biais [de promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) est masquée par le créateur du script. Cette conversion ne prend pas en charge les [types d’union](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) qui incluent des `ExcelScript` types et des types définis par l’utilisateur. Dans ce cas, le script est renvoyé au script, mais le compilateur de script Office ne l’attend pas et le créateur du script ne peut pas interagir avec `Promise` le `Promise` .

L’exemple de code suivant montre une union non prise en service entre `ExcelScript.Table` et une `MyTable` interface personnalisée.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getActiveWorksheet();

  // This union is not supported.
  const tableOrMyTable: ExcelScript.Table | MyTable = selectedSheet.getTables()[0];

  // `getName` returns a promise that can't be resolved by the script.
  const name = tableOrMyTable.getName();

  // This logs "{}" instead of the table name.
  console.log(name);
}

interface MyTable {
  getName(): string
}
```

## <a name="performance-warnings"></a>Avertissements de performances

Le [linter](https://wikipedia.org/wiki/Lint_(software)) de l’éditeur de code avertit si le script peut avoir des problèmes de performances. Les cas et la façon de les contourner sont documentés dans Améliorer les performances de [vos scripts Office.](web-client-performance.md)

## <a name="external-api-calls"></a>Appels d’API externes

Pour plus d’informations, voir la prise en charge des appels [d’API externes dans Office Scripts.](external-calls.md)

## <a name="see-also"></a>Voir aussi

* [Principes de base pour la rédaction de scripts Office en Excel sur le web](scripting-fundamentals.md)
* [Améliorer les performances de vos scripts Office de gestion](web-client-performance.md)
