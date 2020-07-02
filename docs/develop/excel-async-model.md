---
title: Utilisation des API Async de scripts Office pour prendre en charge les scripts hérités
description: Introduction sur les API Async Office scripts et utilisation du modèle Load/Sync pour les scripts hérités.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 78a09232060d862a4e0944356ba2f33f7a264ea1
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999284"
---
# <a name="using-the-office-scripts-async-apis-to-support-legacy-scripts"></a>Utilisation des API Async de scripts Office pour prendre en charge les scripts hérités

Cet article vous apprend à écrire des scripts à l’aide des API héritées, async,. Ces API ont les mêmes fonctionnalités de base que les API de scripts Office, synchrones, mais elles exigent que votre script contrôle la synchronisation des données entre le script et le classeur.

> [!IMPORTANT]
> Le modèle Async ne peut être utilisé qu’avec des scripts créés avant l’implémentation du [modèle d’API](scripting-fundamentals.md?view=office-scripts)actuel. Les scripts sont définitivement verrouillés sur le modèle d’API qu’ils ont lors de leur création. Cela signifie également que si vous souhaitez convertir un script hérité en un nouveau modèle, vous devez utiliser un nouveau script. Nous vous recommandons de mettre à jour vos anciens scripts vers le nouveau modèle lorsque vous effectuez des modifications, étant donné que le modèle actuel est plus facile à utiliser. La rubrique [conversion de scripts Async hérités en la section de modèle actuel](#converting-legacy-async-scripts-to-the-current-model) comporte des conseils sur la façon d’effectuer cette transition.

## <a name="main-function"></a>Fonction `main` :

Les scripts qui utilisent les API Async ont une `main` fonction différente. Il s’agit d’une `async` fonction qui a `Excel.RequestContext` comme premier paramètre.

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>Contexte

La fonction `main` accepte un paramètre `Excel.RequestContext`, nommé `context`. Vous devez imaginer le `context` comme un pont entre le script et le classeur. Le script accède au classeur avec l’objet `context` et utilise ce `context` pour envoyer et recevoir des données.

L’objet `context` est nécessaire car le script et Excel sont exécutés dans différents processus et emplacements. Le script doit apporter des modifications ou rechercher les données du classeur dans le cloud. L’objet `context` gère ces opérations.

## <a name="sync-and-load"></a>Synchronisation et chargement

Comme le script et le classeur s’exécutent dans des emplacements différents, le transfert de données entre les deux prend du temps. Dans l’API Async, les commandes sont mises en file d’attente jusqu’à ce que le script appelle explicitement l' `sync` opération pour synchroniser le script et le classeur. Le script peut fonctionner de façon indépendante jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :

- Lisez les données du classeur (en suivant une `load`opération de ou une méthode qui renvoie une [ClientResult](/javascript/api/office-scripts/excelscript/excel.clientresult?view=office-scripts-async)).
- Écrire les données dans le classeur (généralement quand le script est terminé).

L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :

![Un diagramme montrant les opérations de lecture et d’écriture effectuées dans le classeur à partir du script.](../images/load-sync.png)

### <a name="sync"></a>Synchronisation

Chaque fois que votre script Async doit lire ou écrire des données dans le classeur, appelez la `RequestContext.sync` méthode comme illustré ci-dessous :

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` est appelé implicitement à la fin d’un script.

Une fois l’opération `sync` terminée, le classeur se met à jour pour illustrer les opérations d’écriture que le script a spécifiées. Une opération d’écriture définit une propriété sur un objet Excel (par exemple : `range.format.fill.color = "red"`) ou appelle une méthode qui modifie une propriété (par exemple : `range.format.autoFitColumns()`). L’opération `sync` lit également les valeurs du classeur demandées par le script à l’aide d’une opération `load` ou d’une méthode renvoyant une `ClientResult` (comme indiqué dans la section suivante).

La synchronisation du script avec le classeur peut prendre du temps, en fonction de votre réseau. Réduisez le nombre d' `sync` appels pour aider votre script à s’exécuter rapidement. Dans le cas contraire, les API asynchrones ne sont pas plus rapides que les API synchrones standard.

### <a name="load"></a>Charger

Un script Async doit charger les données du classeur avant de le lire. Toutefois, le chargement des données à partir de l’intégralité du classeur réduirait considérablement la vitesse du script. La `load` méthode permet à votre script d’indiquer spécifiquement quelles données doivent être récupérées à partir du classeur.

La méthode `load` est disponible sur tous les objets Excel. Le script doit charger les propriétés d’un objet avant de pouvoir les lire. Si ce n’est pas le cas, une erreur est générée.

Les exemples suivants utilisent un objet `Range` pour illustrer les trois méthodes utilisées par `load` pour charger les données.

|Objectif |Exemple de commande | Effet |
|:--|:--|:--|
|Charger une propriété |`myRange.load("values");` | Charge une seule propriété. Dans ce cas, le tableau à deux dimensions des valeurs dans cette plage. |
|Charger plusieurs propriétés |`myRange.load("values, rowCount, columnCount");`| Charge toutes les propriétés d’une liste, qui sont délimitées par des virgules. Dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes. |
|Tout charger | `myRange.load();`|Charge toutes les propriétés de la plage. Il ne s’agit pas d’une solution recommandée, car elle ralentit votre script en obtenant des données inutiles. Ne l’utilisez que si vous testez votre script ou si vous avez besoin de chaque propriété de l’objet. |

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

Vous pouvez également charger des propriétés sur l’ensemble d’une collection. Chaque objet collection de l’API Async a une `items` propriété qui est un tableau contenant les objets de cette collection. L’utilisation de `items` comme point de départ d’un appel hiérarchique (`items\myProperty`) pour que `load` charge les propriétés spécifiées sur chacun de ces éléments. L’exemple suivant charge la propriété `resolved` sur tous les objets `Comment` dans l’objet `CommentCollection` d’une feuille de calcul.

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

Les méthodes de l’API Async qui renvoient des informations à partir du classeur ont un modèle similaire pour le `load` / `sync` paradigme. Par exemple, `TableCollection.getCount` obtient le nombre de tableaux dans la collection. `getCount` renvoie une `ClientResult<number>`, ce qui signifie que la propriété `value` dans le renvoie `ClientResult` est un nombre. Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé. À l’instar du chargement d’une propriété, la valeur `value` est une valeur « vide » locale jusqu’à cet appel`sync`.

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

## <a name="converting-legacy-async-scripts-to-the-current-model"></a>Conversion de scripts Async hérités en modèle actuel

Le modèle d’API actuel n’utilise pas `load` , `sync` , ni un `RequestContext` . Les scripts sont ainsi beaucoup plus faciles à écrire et à gérer. La meilleure ressource pour convertir les anciens scripts est le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts). Dans ce cas, vous pouvez demander de l’aide à la communauté pour des scénarios spécifiques. Les conseils suivants devraient vous aider à décrire les étapes générales à suivre.

1. Créez un script et copiez-y l’ancien code Async. Veillez à ne pas inclure l’ancienne `main` signature de méthode, en utilisant la version actuelle à la `function main(workbook: ExcelScript.Workbook)` place.

2. Supprimez tous `load` les `sync` appels et. Ils ne sont plus nécessaires.

3. Toutes les propriétés ont été supprimées. À présent, vous accédez à ces objets par le biais `get` de et de `set` méthodes, vous devrez donc changer ces références de propriété en appels de méthode. Par exemple, au lieu de définir la couleur de remplissage d’une cellule par le biais d’un accès aux propriétés comme suit : `mySheet.getRange("A2:C2").format.fill.color = "blue";` , vous utilisez des méthodes comme celle-ci :`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. Les classes de collection ont été remplacées par des tableaux. Les `add` `get` méthodes et de ces classes de collection ont été déplacées vers l’objet propriétaire de la collection, de sorte que vos références doivent être mises à jour en conséquence. Par exemple, pour obtenir un graphique nommé « MyChart » à partir de la première feuille de calcul du classeur, utilisez le code suivant : `workbook.getWorksheets()[0].getChart("MyChart");` . Notez le `[0]` pour accéder à la première valeur de la `Worksheet[]` renvoyée par `getWorksheets()` .

5. Certaines méthodes ont été renommées pour des raisons de clarté et de commodité. Pour plus d’informations, consultez la référence de l' [API scripts Office](/javascript/api/office-scripts/overview?view=office-scripts) .

## <a name="office-scripts-async-api-reference-documentation"></a>Documentation de référence de l’API asynchrone de scripts Office

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
