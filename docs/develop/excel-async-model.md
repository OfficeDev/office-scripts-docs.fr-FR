---
title: Soutenez les scripts Office plus anciens qui utilisent les API async
description: Une amorce sur les Office scripts Async API et comment utiliser le modèle de charge / synchronisation pour les scripts plus anciens.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 80a1c0dec5393d8882ddb37eea5f81ef23b1ebb1
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545074"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Soutenez les scripts Office plus anciens qui utilisent les API async

Cet article vous enseigne comment maintenir et mettre à jour les scripts qui utilisent les API async de l’ancien modèle. Ces API ont les mêmes fonctionnalités de base que les API scripts Office synchrones désormais standard, mais elles nécessitent que votre script contrôle la synchronisation des données entre le script et le cahier de travail.

> [!IMPORTANT]
> Le modèle async ne peut être utilisé qu’avec des scripts créés avant la mise en œuvre du modèle [API actuel.](scripting-fundamentals.md) Les scripts sont verrouillés en permanence sur le modèle API qu’ils ont lors de la création. Cela signifie également que si vous souhaitez convertir un ancien script au nouveau modèle, vous devez créer un tout nouveau script. Nous vous recommandons de mettre à jour vos anciens scripts sur le nouveau modèle lors des modifications, puisque le modèle actuel est plus facile à utiliser. Les [scripts async de conversion à la](#convert-async-scripts-to-the-current-model) section modèle actuelle a des conseils sur la façon de faire cette transition.

## <a name="older-main-function-signature"></a>Signature `main` de fonction plus ancienne

Les scripts qui utilisent les API async ont une fonction `main` différente. C’est une `async` fonction qui a un comme premier `Excel.RequestContext` paramètre.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Contexte

La fonction `main` accepte un paramètre `Excel.RequestContext`, nommé `context`. Vous devez imaginer le `context` comme un pont entre le script et le classeur. Le script accède au classeur avec l’objet `context` et utilise ce `context` pour envoyer et recevoir des données.

L’objet `context` est nécessaire car le script et Excel sont exécutés dans différents processus et emplacements. Le script doit apporter des modifications ou rechercher les données du classeur dans le cloud. L’objet `context` gère ces opérations.

## <a name="sync-and-load"></a>Synchronisation et charge

Comme le script et le classeur s’exécutent dans des emplacements différents, le transfert de données entre les deux prend du temps. Dans l’API async, les commandes sont en file d’attente jusqu’à ce que le script `sync` appelle explicitement l’opération pour synchroniser le script et le cahier de travail. Le script peut fonctionner de façon indépendante jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :

- Lisez les données du classeur (en suivant une `load`opération de ou une méthode qui renvoie une [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Écrire les données dans le classeur (généralement quand le script est terminé).

L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :

:::image type="content" source="../images/load-sync.png" alt-text="Un diagramme montrant les opérations de lecture et d’écriture allant au cahier de travail à partir du script":::

### <a name="sync"></a>Synchronisation

Chaque fois que votre script async doit lire des données ou écrire des données sur le cahier de travail, appelez la méthode comme `RequestContext.sync` indiqué dans l’extrait de code suivant :

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` est appelé implicitement à la fin d’un script.

Une fois l’opération `sync` terminée, le classeur se met à jour pour illustrer les opérations d’écriture que le script a spécifiées. Une opération d’écriture consiste à définir n’importe quelle propriété sur un objet Excel (p. ex., ) ou à `range.format.fill.color = "red"` appeler une méthode qui modifie une propriété (p. ex., `range.format.autoFitColumns()` ). L’opération `sync` lit également les valeurs du classeur demandées par le script à l’aide d’une opération `load` ou d’une méthode renvoyant une `ClientResult` (comme indiqué dans la section suivante).

La synchronisation du script avec le classeur peut prendre du temps, en fonction de votre réseau. Réduisez au minimum le nombre `sync` d’appels pour aider votre script à s’exécuter rapidement. Sinon, les API async ne sont pas plus rapides que les API synchrones standard.

### <a name="load"></a>Charger

Un script async doit charger les données du cahier de travail avant de les lire. Toutefois, le chargement des données de l’ensemble du manuel réduirait considérablement la vitesse du script. La `load` méthode permet à votre script d’indiquer spécifiquement quelles données doivent être récupérées à partir du cahier de travail.

La méthode `load` est disponible sur tous les objets Excel. Le script doit charger les propriétés d’un objet avant de pouvoir les lire. Ne pas le faire entraîne une erreur.

Les exemples suivants utilisent un objet `Range` pour illustrer les trois méthodes utilisées par `load` pour charger les données.

|Objectif |Exemple de commande | Effet |
|:--|:--|:--|
|Charger une propriété |`myRange.load("values");` | Charge une seule propriété. Dans ce cas, le tableau à deux dimensions des valeurs dans cette plage. |
|Charger plusieurs propriétés |`myRange.load("values, rowCount, columnCount");`| Charge toutes les propriétés d’une liste, qui sont délimitées par des virgules. Dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes. |
|Tout charger | `myRange.load();`|Charge toutes les propriétés de la plage. Ce n’est pas une solution recommandée, car il va ralentir votre script en obtenant des données inutiles. Utilisez-le uniquement lors de la mise à l’essai de votre script ou si vous avez besoin de chaque propriété de l’objet. |

Le script doit appeler `context.sync()` avant de lire les valeurs chargées.

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

Vous pouvez également charger des propriétés sur l’ensemble d’une collection. Chaque objet de collection de l’API async possède `items` une propriété qui est un tableau contenant les objets de cette collection. L’utilisation de `items` comme point de départ d’un appel hiérarchique (`items\myProperty`) pour que `load` charge les propriétés spécifiées sur chacun de ces éléments. L’exemple suivant charge la propriété `resolved` sur tous les objets `Comment` dans l’objet `CommentCollection` d’une feuille de calcul.

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a>ClientResult

Les méthodes de l’API async qui retournent les informations du cahier de travail ont un modèle similaire au `load` / `sync` paradigme. Par exemple, `TableCollection.getCount` obtient le nombre de tableaux dans la collection. `getCount` renvoie un `ClientResult<number>`, ce qui signifie que la propriété `value` dans le [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) renvoyé est un nombre. Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé. À l’instar du chargement d’une propriété, la valeur `value` est une valeur « vide » locale jusqu’à cet appel`sync`.

Le script suivant fournit le nombre total de tableaux dans le classeur et enregistre ce nombre sur la console.

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="convert-async-scripts-to-the-current-model"></a>Convertir les scripts async au modèle actuel

Le modèle API actuel n’utilise `load` pas , ou un `sync` `RequestContext` . Cela rend les scripts beaucoup plus faciles à écrire et à entretenir. Votre meilleure ressource pour convertir de vieux scripts est [Microsoft Q&A](/answers/topics/office-scripts-dev.html). Là, vous pouvez demander de l’aide à la communauté avec des scénarios spécifiques. Les conseils suivants devraient vous aider à décrire les mesures générales que vous devrez prendre.

1. Créez un nouveau script et copiez l’ancien code async en lui. Assurez-vous de ne pas inclure l’ancienne `main` signature de la méthode, en utilisant le courant à `function main(workbook: ExcelScript.Workbook)` la place.

2. Retirez tous les `load` appels et `sync` les appels. Ils ne sont plus nécessaires.

3. Toutes les propriétés ont été supprimées. Vous accédez maintenant à ces `get` objets `set` et méthodes, de sorte que vous aurez besoin de passer ces références de propriété à des appels de méthode. Par exemple, au lieu de définir la couleur de remplissage d’une cellule par l’accès à la propriété comme `mySheet.getRange("A2:C2").format.fill.color = "blue";` ceci: , vous allez maintenant utiliser des méthodes comme celle-ci: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Les classes de collection ont été remplacées par des tableaux. Les `add` méthodes de ces classes de collecte ont été déplacées vers `get` l’objet qui possédait la collection, de sorte que vos références doivent être mises à jour en conséquence. Par exemple, pour obtenir un tableau nommé « MyChart » à partir de la première feuille de travail dans le cahier de travail, utilisez le code suivant: `workbook.getWorksheets()[0].getChart("MyChart");` . Notez `[0]` l’accès à la première valeur de la `Worksheet[]` retournée par `getWorksheets()` .

5. Certaines méthodes ont été renommées pour plus de clarté et ajoutées pour plus de commodité. Veuillez consulter la référence [Office’API scripts pour plus](/javascript/api/office-scripts/overview) de détails.

## <a name="office-scripts-async-api-reference-documentation"></a>Office Scripts async API documentation de référence

Les API async sont équivalentes à celles utilisées dans Office Add-ins. La documentation de référence se trouve [dans la section Excel de la référence Office’API JavaScript add-ins](/javascript/api/excel?view=excel-js-online&preserve-view=true).
