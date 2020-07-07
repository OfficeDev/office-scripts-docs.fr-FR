---
title: Utilisation des API Async de scripts Office pour prendre en charge les scripts hérités
description: Introduction sur les API Async Office scripts et utilisation du modèle Load/Sync pour les scripts hérités.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 6c31a39c8e1fe53f2f5587183a6b32e100d2b457
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043397"
---
# <a name="using-the-office-scripts-async-apis-to-support-legacy-scripts"></a><span data-ttu-id="31e0a-103">Utilisation des API Async de scripts Office pour prendre en charge les scripts hérités</span><span class="sxs-lookup"><span data-stu-id="31e0a-103">Using the Office Scripts Async APIs to support legacy scripts</span></span>

<span data-ttu-id="31e0a-104">Cet article vous apprend à écrire des scripts à l’aide des API héritées, async,.</span><span class="sxs-lookup"><span data-stu-id="31e0a-104">This article will teach you how to write scripts using the legacy, async, APIs.</span></span> <span data-ttu-id="31e0a-105">Ces API ont les mêmes fonctionnalités de base que les API de scripts Office, synchrones, mais elles exigent que votre script contrôle la synchronisation des données entre le script et le classeur.</span><span class="sxs-lookup"><span data-stu-id="31e0a-105">These APIs have the same core functionality as the standard, synchronous Office Scripts APIs, but they require that your script control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="31e0a-106">Le modèle Async ne peut être utilisé qu’avec des scripts créés avant l’implémentation du [modèle d’API](scripting-fundamentals.md?view=office-scripts)actuel.</span><span class="sxs-lookup"><span data-stu-id="31e0a-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md?view=office-scripts).</span></span> <span data-ttu-id="31e0a-107">Les scripts sont définitivement verrouillés sur le modèle d’API qu’ils ont lors de leur création.</span><span class="sxs-lookup"><span data-stu-id="31e0a-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="31e0a-108">Cela signifie également que si vous souhaitez convertir un script hérité en un nouveau modèle, vous devez utiliser un nouveau script.</span><span class="sxs-lookup"><span data-stu-id="31e0a-108">This also means that if you want to convert a legacy script to the new model, you must use a brand new script.</span></span> <span data-ttu-id="31e0a-109">Nous vous recommandons de mettre à jour vos anciens scripts vers le nouveau modèle lorsque vous effectuez des modifications, étant donné que le modèle actuel est plus facile à utiliser.</span><span class="sxs-lookup"><span data-stu-id="31e0a-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="31e0a-110">La rubrique [conversion de scripts Async hérités en la section de modèle actuel](#converting-legacy-async-scripts-to-the-current-model) comporte des conseils sur la façon d’effectuer cette transition.</span><span class="sxs-lookup"><span data-stu-id="31e0a-110">The [Converting legacy async scripts to the current model](#converting-legacy-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="31e0a-111">Fonction `main` :</span><span class="sxs-lookup"><span data-stu-id="31e0a-111">`main` function</span></span>

<span data-ttu-id="31e0a-112">Les scripts qui utilisent les API Async ont une `main` fonction différente.</span><span class="sxs-lookup"><span data-stu-id="31e0a-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="31e0a-113">Il s’agit d’une `async` fonction qui a `Excel.RequestContext` comme premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="31e0a-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="31e0a-114">Contexte</span><span class="sxs-lookup"><span data-stu-id="31e0a-114">Context</span></span>

<span data-ttu-id="31e0a-115">La fonction `main` accepte un paramètre `Excel.RequestContext`, nommé `context`.</span><span class="sxs-lookup"><span data-stu-id="31e0a-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="31e0a-116">Vous devez imaginer le `context` comme un pont entre le script et le classeur.</span><span class="sxs-lookup"><span data-stu-id="31e0a-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="31e0a-117">Le script accède au classeur avec l’objet `context` et utilise ce `context` pour envoyer et recevoir des données.</span><span class="sxs-lookup"><span data-stu-id="31e0a-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="31e0a-118">L’objet `context` est nécessaire car le script et Excel sont exécutés dans différents processus et emplacements.</span><span class="sxs-lookup"><span data-stu-id="31e0a-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="31e0a-119">Le script doit apporter des modifications ou rechercher les données du classeur dans le cloud.</span><span class="sxs-lookup"><span data-stu-id="31e0a-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="31e0a-120">L’objet `context` gère ces opérations.</span><span class="sxs-lookup"><span data-stu-id="31e0a-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="31e0a-121">Synchronisation et chargement</span><span class="sxs-lookup"><span data-stu-id="31e0a-121">Sync and Load</span></span>

<span data-ttu-id="31e0a-122">Comme le script et le classeur s’exécutent dans des emplacements différents, le transfert de données entre les deux prend du temps.</span><span class="sxs-lookup"><span data-stu-id="31e0a-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="31e0a-123">Dans l’API Async, les commandes sont mises en file d’attente jusqu’à ce que le script appelle explicitement l' `sync` opération pour synchroniser le script et le classeur.</span><span class="sxs-lookup"><span data-stu-id="31e0a-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="31e0a-124">Le script peut fonctionner de façon indépendante jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="31e0a-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="31e0a-125">Lisez les données du classeur (en suivant une `load`opération de ou une méthode qui renvoie une [ClientResult](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async)).</span><span class="sxs-lookup"><span data-stu-id="31e0a-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async)).</span></span>
- <span data-ttu-id="31e0a-126">Écrire les données dans le classeur (généralement quand le script est terminé).</span><span class="sxs-lookup"><span data-stu-id="31e0a-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="31e0a-127">L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :</span><span class="sxs-lookup"><span data-stu-id="31e0a-127">The following image shows an example control flow between the script and workbook:</span></span>

![Un diagramme montrant les opérations de lecture et d’écriture effectuées dans le classeur à partir du script.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="31e0a-129">Synchronisation</span><span class="sxs-lookup"><span data-stu-id="31e0a-129">Sync</span></span>

<span data-ttu-id="31e0a-130">Chaque fois que votre script Async doit lire ou écrire des données dans le classeur, appelez la `RequestContext.sync` méthode comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="31e0a-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="31e0a-131">`context.sync()` est appelé implicitement à la fin d’un script.</span><span class="sxs-lookup"><span data-stu-id="31e0a-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="31e0a-132">Une fois l’opération `sync` terminée, le classeur se met à jour pour illustrer les opérations d’écriture que le script a spécifiées.</span><span class="sxs-lookup"><span data-stu-id="31e0a-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="31e0a-133">Une opération d’écriture définit une propriété sur un objet Excel (par exemple : `range.format.fill.color = "red"`) ou appelle une méthode qui modifie une propriété (par exemple : `range.format.autoFitColumns()`).</span><span class="sxs-lookup"><span data-stu-id="31e0a-133">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="31e0a-134">L’opération `sync` lit également les valeurs du classeur demandées par le script à l’aide d’une opération `load` ou d’une méthode renvoyant une `ClientResult` (comme indiqué dans la section suivante).</span><span class="sxs-lookup"><span data-stu-id="31e0a-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="31e0a-135">La synchronisation du script avec le classeur peut prendre du temps, en fonction de votre réseau.</span><span class="sxs-lookup"><span data-stu-id="31e0a-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="31e0a-136">Réduisez le nombre d' `sync` appels pour aider votre script à s’exécuter rapidement.</span><span class="sxs-lookup"><span data-stu-id="31e0a-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="31e0a-137">Dans le cas contraire, les API asynchrones ne sont pas plus rapides que les API synchrones standard.</span><span class="sxs-lookup"><span data-stu-id="31e0a-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="31e0a-138">Charger</span><span class="sxs-lookup"><span data-stu-id="31e0a-138">Load</span></span>

<span data-ttu-id="31e0a-139">Un script Async doit charger les données du classeur avant de le lire.</span><span class="sxs-lookup"><span data-stu-id="31e0a-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="31e0a-140">Toutefois, le chargement des données à partir de l’intégralité du classeur réduirait considérablement la vitesse du script.</span><span class="sxs-lookup"><span data-stu-id="31e0a-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="31e0a-141">La `load` méthode permet à votre script d’indiquer spécifiquement quelles données doivent être récupérées à partir du classeur.</span><span class="sxs-lookup"><span data-stu-id="31e0a-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="31e0a-142">La méthode `load` est disponible sur tous les objets Excel.</span><span class="sxs-lookup"><span data-stu-id="31e0a-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="31e0a-143">Le script doit charger les propriétés d’un objet avant de pouvoir les lire.</span><span class="sxs-lookup"><span data-stu-id="31e0a-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="31e0a-144">Si ce n’est pas le cas, une erreur est générée.</span><span class="sxs-lookup"><span data-stu-id="31e0a-144">Not doing so results in an error.</span></span>

<span data-ttu-id="31e0a-145">Les exemples suivants utilisent un objet `Range` pour illustrer les trois méthodes utilisées par `load` pour charger les données.</span><span class="sxs-lookup"><span data-stu-id="31e0a-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="31e0a-146">Objectif</span><span class="sxs-lookup"><span data-stu-id="31e0a-146">Intent</span></span> |<span data-ttu-id="31e0a-147">Exemple de commande</span><span class="sxs-lookup"><span data-stu-id="31e0a-147">Example Command</span></span> | <span data-ttu-id="31e0a-148">Effet</span><span class="sxs-lookup"><span data-stu-id="31e0a-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="31e0a-149">Charger une propriété</span><span class="sxs-lookup"><span data-stu-id="31e0a-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="31e0a-150">Charge une seule propriété. Dans ce cas, le tableau à deux dimensions des valeurs dans cette plage.</span><span class="sxs-lookup"><span data-stu-id="31e0a-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="31e0a-151">Charger plusieurs propriétés</span><span class="sxs-lookup"><span data-stu-id="31e0a-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="31e0a-152">Charge toutes les propriétés d’une liste, qui sont délimitées par des virgules. Dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes.</span><span class="sxs-lookup"><span data-stu-id="31e0a-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="31e0a-153">Tout charger</span><span class="sxs-lookup"><span data-stu-id="31e0a-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="31e0a-154">Charge toutes les propriétés de la plage.</span><span class="sxs-lookup"><span data-stu-id="31e0a-154">Loads all the properties on the range.</span></span> <span data-ttu-id="31e0a-155">Il ne s’agit pas d’une solution recommandée, car elle ralentit votre script en obtenant des données inutiles.</span><span class="sxs-lookup"><span data-stu-id="31e0a-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="31e0a-156">Ne l’utilisez que si vous testez votre script ou si vous avez besoin de chaque propriété de l’objet.</span><span class="sxs-lookup"><span data-stu-id="31e0a-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="31e0a-157">Le script doit appeler `context.sync()` avant de lire les valeurs chargées.</span><span class="sxs-lookup"><span data-stu-id="31e0a-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

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

<span data-ttu-id="31e0a-158">Vous pouvez également charger des propriétés sur l’ensemble d’une collection.</span><span class="sxs-lookup"><span data-stu-id="31e0a-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="31e0a-159">Chaque objet collection de l’API Async a une `items` propriété qui est un tableau contenant les objets de cette collection.</span><span class="sxs-lookup"><span data-stu-id="31e0a-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="31e0a-160">L’utilisation de `items` comme point de départ d’un appel hiérarchique (`items\myProperty`) pour que `load` charge les propriétés spécifiées sur chacun de ces éléments.</span><span class="sxs-lookup"><span data-stu-id="31e0a-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="31e0a-161">L’exemple suivant charge la propriété `resolved` sur tous les objets `Comment` dans l’objet `CommentCollection` d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="31e0a-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

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

### <a name="clientresult"></a><span data-ttu-id="31e0a-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="31e0a-162">ClientResult</span></span>

<span data-ttu-id="31e0a-163">Les méthodes de l’API Async qui renvoient des informations à partir du classeur ont un modèle similaire pour le `load` / `sync` paradigme.</span><span class="sxs-lookup"><span data-stu-id="31e0a-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="31e0a-164">Par exemple, `TableCollection.getCount` obtient le nombre de tableaux dans la collection.</span><span class="sxs-lookup"><span data-stu-id="31e0a-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="31e0a-165">`getCount`renvoie un `ClientResult<number>` , ce qui signifie que la `value` propriété dans le renvoyé [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) est un nombre.</span><span class="sxs-lookup"><span data-stu-id="31e0a-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) is a number.</span></span> <span data-ttu-id="31e0a-166">Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé.</span><span class="sxs-lookup"><span data-stu-id="31e0a-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="31e0a-167">À l’instar du chargement d’une propriété, la valeur `value` est une valeur « vide » locale jusqu’à cet appel`sync`.</span><span class="sxs-lookup"><span data-stu-id="31e0a-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="31e0a-168">Le script suivant fournit le nombre total de tableaux dans le classeur et enregistre ce nombre sur la console.</span><span class="sxs-lookup"><span data-stu-id="31e0a-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

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

## <a name="converting-legacy-async-scripts-to-the-current-model"></a><span data-ttu-id="31e0a-169">Conversion de scripts Async hérités en modèle actuel</span><span class="sxs-lookup"><span data-stu-id="31e0a-169">Converting legacy async scripts to the current model</span></span>

<span data-ttu-id="31e0a-170">Le modèle d’API actuel n’utilise pas `load` , `sync` , ni un `RequestContext` .</span><span class="sxs-lookup"><span data-stu-id="31e0a-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="31e0a-171">Les scripts sont ainsi beaucoup plus faciles à écrire et à gérer.</span><span class="sxs-lookup"><span data-stu-id="31e0a-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="31e0a-172">La meilleure ressource pour convertir les anciens scripts est le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="31e0a-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="31e0a-173">Dans ce cas, vous pouvez demander de l’aide à la communauté pour des scénarios spécifiques.</span><span class="sxs-lookup"><span data-stu-id="31e0a-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="31e0a-174">Les conseils suivants devraient vous aider à décrire les étapes générales à suivre.</span><span class="sxs-lookup"><span data-stu-id="31e0a-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="31e0a-175">Créez un script et copiez-y l’ancien code Async.</span><span class="sxs-lookup"><span data-stu-id="31e0a-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="31e0a-176">Veillez à ne pas inclure l’ancienne `main` signature de méthode, en utilisant la version actuelle à la `function main(workbook: ExcelScript.Workbook)` place.</span><span class="sxs-lookup"><span data-stu-id="31e0a-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="31e0a-177">Supprimez tous `load` les `sync` appels et.</span><span class="sxs-lookup"><span data-stu-id="31e0a-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="31e0a-178">Ils ne sont plus nécessaires.</span><span class="sxs-lookup"><span data-stu-id="31e0a-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="31e0a-179">Toutes les propriétés ont été supprimées.</span><span class="sxs-lookup"><span data-stu-id="31e0a-179">All properties have been removed.</span></span> <span data-ttu-id="31e0a-180">À présent, vous accédez à ces objets par le biais `get` de et de `set` méthodes, vous devrez donc changer ces références de propriété en appels de méthode.</span><span class="sxs-lookup"><span data-stu-id="31e0a-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="31e0a-181">Par exemple, au lieu de définir la couleur de remplissage d’une cellule par le biais d’un accès aux propriétés comme suit : `mySheet.getRange("A2:C2").format.fill.color = "blue";` , vous utilisez des méthodes comme celle-ci :`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="31e0a-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="31e0a-182">Les classes de collection ont été remplacées par des tableaux.</span><span class="sxs-lookup"><span data-stu-id="31e0a-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="31e0a-183">Les `add` `get` méthodes et de ces classes de collection ont été déplacées vers l’objet propriétaire de la collection, de sorte que vos références doivent être mises à jour en conséquence.</span><span class="sxs-lookup"><span data-stu-id="31e0a-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="31e0a-184">Par exemple, pour obtenir un graphique nommé « MyChart » à partir de la première feuille de calcul du classeur, utilisez le code suivant : `workbook.getWorksheets()[0].getChart("MyChart");` .</span><span class="sxs-lookup"><span data-stu-id="31e0a-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="31e0a-185">Notez le `[0]` pour accéder à la première valeur de la `Worksheet[]` renvoyée par `getWorksheets()` .</span><span class="sxs-lookup"><span data-stu-id="31e0a-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="31e0a-186">Certaines méthodes ont été renommées pour des raisons de clarté et de commodité.</span><span class="sxs-lookup"><span data-stu-id="31e0a-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="31e0a-187">Pour plus d’informations, consultez la référence de l' [API scripts Office](/javascript/api/office-scripts/overview?view=office-scripts) .</span><span class="sxs-lookup"><span data-stu-id="31e0a-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview?view=office-scripts) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="31e0a-188">Documentation de référence de l’API asynchrone de scripts Office</span><span class="sxs-lookup"><span data-stu-id="31e0a-188">Office Scripts Async API reference documentation</span></span>

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
