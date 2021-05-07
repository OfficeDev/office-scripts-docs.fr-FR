---
title: Prise en charge Office scripts plus anciens qui utilisent les API async
description: A primer on the Office Scripts Async APIs and how to use the load/sync pattern for older scripts.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: 437fb2e389d6d3963f93cdb44c5529749c4d3569
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232409"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>Prise en charge Office scripts plus anciens qui utilisent les API async

Cet article vous montre comment gérer et mettre à jour des scripts qui utilisent les API async de l’ancien modèle. Ces API ont les mêmes fonctionnalités de base que les API Office Scripts désormais standard et synchrones, mais elles nécessitent votre script pour contrôler la synchronisation des données entre le script et le workbook.

> [!IMPORTANT]
> Le modèle async ne peut être utilisé qu’avec des scripts créés avant l’implémentation du modèle [API actuel.](scripting-fundamentals.md) Les scripts sont définitivement verrouillés sur le modèle d’API dont ils ont lors de leur création. Cela signifie également que si vous souhaitez convertir un ancien script vers le nouveau modèle, vous devez créer un tout nouveau script. Nous vous recommandons de mettre à jour vos anciens scripts vers le nouveau modèle lorsque vous a apporté des modifications, car le modèle actuel est plus facile à utiliser. La section [Conversion de scripts async](#converting-async-scripts-to-the-current-model) en modèle actuel contient des conseils sur la façon d’effectuer cette transition.

## <a name="main-function"></a>Fonction `main` :

Les scripts qui utilisent les API async ont une fonction `main` différente. Il s’agit `async` d’une fonction qui a `Excel.RequestContext` un comme premier paramètre.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Contexte

La fonction `main` accepte un paramètre `Excel.RequestContext`, nommé `context`. Vous devez imaginer le `context` comme un pont entre le script et le classeur. Le script accède au classeur avec l’objet `context` et utilise ce `context` pour envoyer et recevoir des données.

L’objet `context` est nécessaire car le script et Excel sont exécutés dans différents processus et emplacements. Le script doit apporter des modifications ou rechercher les données du classeur dans le cloud. L’objet `context` gère ces opérations.

## <a name="sync-and-load"></a>Synchronisation et chargement

Comme le script et le classeur s’exécutent dans des emplacements différents, le transfert de données entre les deux prend du temps. Dans l’API async, les commandes sont en file d’attente jusqu’à ce que le script appelle explicitement l’opération pour synchroniser le `sync` script et le workbook. Le script peut fonctionner de façon indépendante jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :

- Lisez les données du classeur (en suivant une `load`opération de ou une méthode qui renvoie une [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).
- Écrire les données dans le classeur (généralement quand le script est terminé).

L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :

:::image type="content" source="../images/load-sync.png" alt-text="Diagramme montrant les opérations de lecture et d’écriture dans le workbook à partir du script":::

### <a name="sync"></a>Synchronisation

Chaque fois que votre script async doit lire ou écrire des données dans le workbook, appelez la méthode `RequestContext.sync` comme illustré ici :

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` est appelé implicitement à la fin d’un script.

Une fois l’opération `sync` terminée, le classeur se met à jour pour illustrer les opérations d’écriture que le script a spécifiées. Une opération d’écriture consiste à définir une propriété sur un objet Excel (par exemple, ) ou à appeler une méthode qui modifie une propriété `range.format.fill.color = "red"` (par exemple, `range.format.autoFitColumns()` ). L’opération `sync` lit également les valeurs du classeur demandées par le script à l’aide d’une opération `load` ou d’une méthode renvoyant une `ClientResult` (comme indiqué dans la section suivante).

La synchronisation du script avec le classeur peut prendre du temps, en fonction de votre réseau. Réduisez le nombre `sync` d’appels pour aider votre script à s’exécuter rapidement. Dans le cas contraire, les API asynchrones ne sont pas plus rapides que les API synchrones standard.

### <a name="load"></a>Charger

Un script async doit charger des données à partir du workbook avant de les lire. Toutefois, le chargement de données à partir de l’intégralité du manuel réduit considérablement la vitesse du script. La `load` méthode permet à votre script d’états spécifiques quelles données doivent être récupérées à partir du workbook.

La méthode `load` est disponible sur tous les objets Excel. Le script doit charger les propriétés d’un objet avant de pouvoir les lire. Si ce n’est pas le cas, une erreur est produite.

Les exemples suivants utilisent un objet `Range` pour illustrer les trois méthodes utilisées par `load` pour charger les données.

|Objectif |Exemple de commande | Effet |
|:--|:--|:--|
|Charger une propriété |`myRange.load("values");` | Charge une seule propriété. Dans ce cas, le tableau à deux dimensions des valeurs dans cette plage. |
|Charger plusieurs propriétés |`myRange.load("values, rowCount, columnCount");`| Charge toutes les propriétés d’une liste, qui sont délimitées par des virgules. Dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes. |
|Tout charger | `myRange.load();`|Charge toutes les propriétés de la plage. Cette solution n’est pas recommandée, car elle ralentit votre script en obtenant des données inutiles. Utilisez-le uniquement lors du test de votre script ou si vous avez besoin de toutes les propriétés de l’objet. |

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

Vous pouvez également charger des propriétés sur l’ensemble d’une collection. Chaque objet de collection dans l’API async possède une propriété qui est un tableau contenant les `items` objets de cette collection. L’utilisation de `items` comme point de départ d’un appel hiérarchique (`items\myProperty`) pour que `load` charge les propriétés spécifiées sur chacun de ces éléments. L’exemple suivant charge la propriété `resolved` sur tous les objets `Comment` dans l’objet `CommentCollection` d’une feuille de calcul.

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

Les méthodes de l’API async qui retournent des informations à partir du manuel ont un modèle similaire au `load` / `sync` paradigme. Par exemple, `TableCollection.getCount` obtient le nombre de tableaux dans la collection. `getCount` renvoie un `ClientResult<number>`, ce qui signifie que la propriété `value` dans le [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) renvoyé est un nombre. Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé. À l’instar du chargement d’une propriété, la valeur `value` est une valeur « vide » locale jusqu’à cet appel`sync`.

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

## <a name="converting-async-scripts-to-the-current-model"></a>Conversion de scripts async en modèle actuel

Le modèle API actuel n’utilise `load` pas , ou un `sync` `RequestContext` . Cela facilite l’écriture et la maintenance des scripts. Votre meilleure ressource pour la conversion d’anciens scripts est [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts). Là, vous pouvez demander de l’aide à la communauté pour des scénarios spécifiques. Les instructions suivantes doivent vous aider à définir les étapes générales à suivre.

1. Créez un script et copiez-y l’ancien code async. Assurez-vous de ne pas inclure l’ancienne `main` signature de méthode en utilisant la signature actuelle à la `function main(workbook: ExcelScript.Workbook)` place.

2. Supprimez tous les `load` `sync` appels. Elles ne sont plus nécessaires.

3. Toutes les propriétés ont été supprimées. Vous accédez maintenant à ces objets par le biais de méthodes, vous devez donc basculer ces références de propriétés vers des appels `get` `set` de méthode. Par exemple, au lieu de définir la couleur de remplissage d’une cellule par le biais de l’accès aux propriétés comme ceci : , vous allez maintenant utiliser des méthodes `mySheet.getRange("A2:C2").format.fill.color = "blue";` comme celle-ci : `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Les classes de collection ont été remplacées par des tableaux. Les méthodes et les méthodes de ces classes de collection ont été déplacées vers l’objet propriétaire de la collection. Vos références doivent donc être mises `add` `get` à jour en conséquence. Par exemple, pour obtenir un graphique nommé « MyChart » à partir de la première feuille de calcul du manuel, utilisez le code suivant `workbook.getWorksheets()[0].getChart("MyChart");` : Notez que `[0]` pour accéder à la première valeur de la valeur `Worksheet[]` renvoyée par `getWorksheets()` .

5. Certaines méthodes ont été renommées pour plus de clarté et ajoutées par souci de commodité. Pour plus [d’informations, Office référence de l’API Scripts.](/javascript/api/office-scripts/overview)

## <a name="office-scripts-async-api-reference-documentation"></a>Office Documentation de référence de l’API async scripts

Les API async sont équivalentes à celles utilisées dans les Office de conférence. La documentation de référence se trouve dans la section Excel de la référence de [l’API JavaScript Office des Office.](/javascript/api/excel?view=excel-js-online&preserve-view=true)
