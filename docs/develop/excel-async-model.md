---
title: Prise en charge d’anciens scripts Office qui utilisent les API async
description: A primer on the Office Scripts Async APIs and how to use the load/sync pattern for older scripts.
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: be7847efe59dc6026875b8a8e3b3c93e0eb82e4d
ms.sourcegitcommit: 345f1dd96d80471b246044b199fe11126a192a88
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50242024"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a><span data-ttu-id="44bef-103">Prise en charge d’anciens scripts Office qui utilisent les API async</span><span class="sxs-lookup"><span data-stu-id="44bef-103">Support older Office Scripts that use the async APIs</span></span>

<span data-ttu-id="44bef-104">Cet article vous montre comment gérer et mettre à jour des scripts qui utilisent les API async de l’ancien modèle.</span><span class="sxs-lookup"><span data-stu-id="44bef-104">This article will teach you how to maintain and update scripts that use the older model's async APIs.</span></span> <span data-ttu-id="44bef-105">Ces API ont les mêmes fonctionnalités de base que les API office scripts synchrones désormais standard, mais elles nécessitent votre script pour contrôler la synchronisation des données entre le script et le workbook.</span><span class="sxs-lookup"><span data-stu-id="44bef-105">These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="44bef-106">Le modèle async ne peut être utilisé qu’avec des scripts créés avant l’implémentation du modèle [API actuel.](scripting-fundamentals.md?view=office-scripts&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="44bef-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md?view=office-scripts&preserve-view=true).</span></span> <span data-ttu-id="44bef-107">Les scripts sont définitivement verrouillés sur le modèle d’API dont ils ont besoin lors de leur création.</span><span class="sxs-lookup"><span data-stu-id="44bef-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="44bef-108">Cela signifie également que si vous souhaitez convertir un ancien script vers le nouveau modèle, vous devez créer un tout nouveau script.</span><span class="sxs-lookup"><span data-stu-id="44bef-108">This also means that if you want to convert an old script to the new model, you must create a brand new script.</span></span> <span data-ttu-id="44bef-109">Nous vous recommandons de mettre à jour vos anciens scripts vers le nouveau modèle lorsque vous a apporté des modifications, car le modèle actuel est plus facile à utiliser.</span><span class="sxs-lookup"><span data-stu-id="44bef-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="44bef-110">La section [Conversion de scripts async](#converting-async-scripts-to-the-current-model) en modèle actuel contient des conseils sur la façon d’effectuer cette transition.</span><span class="sxs-lookup"><span data-stu-id="44bef-110">The [Converting async scripts to the current model](#converting-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="44bef-111">Fonction `main` :</span><span class="sxs-lookup"><span data-stu-id="44bef-111">`main` function</span></span>

<span data-ttu-id="44bef-112">Les scripts qui utilisent les API async ont une fonction `main` différente.</span><span class="sxs-lookup"><span data-stu-id="44bef-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="44bef-113">Il s’agit `async` d’une fonction qui a `Excel.RequestContext` un comme premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="44bef-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="44bef-114">Contexte</span><span class="sxs-lookup"><span data-stu-id="44bef-114">Context</span></span>

<span data-ttu-id="44bef-115">La fonction `main` accepte un paramètre `Excel.RequestContext`, nommé `context`.</span><span class="sxs-lookup"><span data-stu-id="44bef-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="44bef-116">Vous devez imaginer le `context` comme un pont entre le script et le classeur.</span><span class="sxs-lookup"><span data-stu-id="44bef-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="44bef-117">Le script accède au classeur avec l’objet `context` et utilise ce `context` pour envoyer et recevoir des données.</span><span class="sxs-lookup"><span data-stu-id="44bef-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="44bef-118">L’objet `context` est nécessaire car le script et Excel sont exécutés dans différents processus et emplacements.</span><span class="sxs-lookup"><span data-stu-id="44bef-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="44bef-119">Le script doit apporter des modifications ou rechercher les données du classeur dans le cloud.</span><span class="sxs-lookup"><span data-stu-id="44bef-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="44bef-120">L’objet `context` gère ces opérations.</span><span class="sxs-lookup"><span data-stu-id="44bef-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="44bef-121">Synchronisation et chargement</span><span class="sxs-lookup"><span data-stu-id="44bef-121">Sync and Load</span></span>

<span data-ttu-id="44bef-122">Comme le script et le classeur s’exécutent dans des emplacements différents, le transfert de données entre les deux prend du temps.</span><span class="sxs-lookup"><span data-stu-id="44bef-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="44bef-123">Dans l’API async, les commandes sont en file d’attente jusqu’à ce que le script appelle explicitement l’opération pour synchroniser le `sync` script et le workbook.</span><span class="sxs-lookup"><span data-stu-id="44bef-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="44bef-124">Le script peut fonctionner de façon indépendante jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="44bef-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="44bef-125">Lisez les données du classeur (en suivant une `load`opération de ou une méthode qui renvoie une [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span><span class="sxs-lookup"><span data-stu-id="44bef-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true)).</span></span>
- <span data-ttu-id="44bef-126">Écrire les données dans le classeur (généralement quand le script est terminé).</span><span class="sxs-lookup"><span data-stu-id="44bef-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="44bef-127">L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :</span><span class="sxs-lookup"><span data-stu-id="44bef-127">The following image shows an example control flow between the script and workbook:</span></span>

![Un diagramme montrant les opérations de lecture et d’écriture effectuées dans le classeur à partir du script.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="44bef-129">Synchronisation</span><span class="sxs-lookup"><span data-stu-id="44bef-129">Sync</span></span>

<span data-ttu-id="44bef-130">Chaque fois que votre script async doit lire ou écrire des données dans le workbook, appelez la méthode `RequestContext.sync` comme indiqué ici :</span><span class="sxs-lookup"><span data-stu-id="44bef-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="44bef-131">`context.sync()` est appelé implicitement à la fin d’un script.</span><span class="sxs-lookup"><span data-stu-id="44bef-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="44bef-132">Une fois l’opération `sync` terminée, le classeur se met à jour pour illustrer les opérations d’écriture que le script a spécifiées.</span><span class="sxs-lookup"><span data-stu-id="44bef-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="44bef-133">Une opération d’écriture consiste à définir une propriété sur un objet Excel (par exemple, ) ou à appeler une méthode qui modifie une propriété `range.format.fill.color = "red"` (par exemple, `range.format.autoFitColumns()` ).</span><span class="sxs-lookup"><span data-stu-id="44bef-133">A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="44bef-134">L’opération `sync` lit également les valeurs du classeur demandées par le script à l’aide d’une opération `load` ou d’une méthode renvoyant une `ClientResult` (comme indiqué dans la section suivante).</span><span class="sxs-lookup"><span data-stu-id="44bef-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="44bef-135">La synchronisation du script avec le classeur peut prendre du temps, en fonction de votre réseau.</span><span class="sxs-lookup"><span data-stu-id="44bef-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="44bef-136">Réduisez le nombre `sync` d’appels pour aider votre script à s’exécuter rapidement.</span><span class="sxs-lookup"><span data-stu-id="44bef-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="44bef-137">Dans le cas contraire, les API asynchrones ne sont pas plus rapides que les API synchrones standard.</span><span class="sxs-lookup"><span data-stu-id="44bef-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="44bef-138">Charger</span><span class="sxs-lookup"><span data-stu-id="44bef-138">Load</span></span>

<span data-ttu-id="44bef-139">Un script async doit charger des données à partir du workbook avant de les lire.</span><span class="sxs-lookup"><span data-stu-id="44bef-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="44bef-140">Toutefois, le chargement des données à partir de l’intégralité du manuel réduit considérablement la vitesse du script.</span><span class="sxs-lookup"><span data-stu-id="44bef-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="44bef-141">La `load` méthode permet à votre script d’états spécifiques quelles données doivent être récupérées à partir du workbook.</span><span class="sxs-lookup"><span data-stu-id="44bef-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="44bef-142">La méthode `load` est disponible sur tous les objets Excel.</span><span class="sxs-lookup"><span data-stu-id="44bef-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="44bef-143">Le script doit charger les propriétés d’un objet avant de pouvoir les lire.</span><span class="sxs-lookup"><span data-stu-id="44bef-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="44bef-144">Si ce n’est pas le cas, une erreur est produite.</span><span class="sxs-lookup"><span data-stu-id="44bef-144">Not doing so results in an error.</span></span>

<span data-ttu-id="44bef-145">Les exemples suivants utilisent un objet `Range` pour illustrer les trois méthodes utilisées par `load` pour charger les données.</span><span class="sxs-lookup"><span data-stu-id="44bef-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="44bef-146">Objectif</span><span class="sxs-lookup"><span data-stu-id="44bef-146">Intent</span></span> |<span data-ttu-id="44bef-147">Exemple de commande</span><span class="sxs-lookup"><span data-stu-id="44bef-147">Example Command</span></span> | <span data-ttu-id="44bef-148">Effet</span><span class="sxs-lookup"><span data-stu-id="44bef-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="44bef-149">Charger une propriété</span><span class="sxs-lookup"><span data-stu-id="44bef-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="44bef-150">Charge une seule propriété. Dans ce cas, le tableau à deux dimensions des valeurs dans cette plage.</span><span class="sxs-lookup"><span data-stu-id="44bef-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="44bef-151">Charger plusieurs propriétés</span><span class="sxs-lookup"><span data-stu-id="44bef-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="44bef-152">Charge toutes les propriétés d’une liste, qui sont délimitées par des virgules. Dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes.</span><span class="sxs-lookup"><span data-stu-id="44bef-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="44bef-153">Tout charger</span><span class="sxs-lookup"><span data-stu-id="44bef-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="44bef-154">Charge toutes les propriétés de la plage.</span><span class="sxs-lookup"><span data-stu-id="44bef-154">Loads all the properties on the range.</span></span> <span data-ttu-id="44bef-155">Cette solution n’est pas recommandée, car elle ralentit votre script en obtenant des données inutiles.</span><span class="sxs-lookup"><span data-stu-id="44bef-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="44bef-156">Utilisez-le uniquement lors du test de votre script ou si vous avez besoin de toutes les propriétés de l’objet.</span><span class="sxs-lookup"><span data-stu-id="44bef-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="44bef-157">Le script doit appeler `context.sync()` avant de lire les valeurs chargées.</span><span class="sxs-lookup"><span data-stu-id="44bef-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

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

<span data-ttu-id="44bef-158">Vous pouvez également charger des propriétés sur l’ensemble d’une collection.</span><span class="sxs-lookup"><span data-stu-id="44bef-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="44bef-159">Chaque objet de collection dans l’API async possède une propriété qui est un tableau contenant les `items` objets de cette collection.</span><span class="sxs-lookup"><span data-stu-id="44bef-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="44bef-160">L’utilisation de `items` comme point de départ d’un appel hiérarchique (`items\myProperty`) pour que `load` charge les propriétés spécifiées sur chacun de ces éléments.</span><span class="sxs-lookup"><span data-stu-id="44bef-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="44bef-161">L’exemple suivant charge la propriété `resolved` sur tous les objets `Comment` dans l’objet `CommentCollection` d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="44bef-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

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

### <a name="clientresult"></a><span data-ttu-id="44bef-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="44bef-162">ClientResult</span></span>

<span data-ttu-id="44bef-163">Les méthodes de l’API async qui retournent des informations à partir du manuel ont un modèle similaire au `load` / `sync` paradigme.</span><span class="sxs-lookup"><span data-stu-id="44bef-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="44bef-164">Par exemple, `TableCollection.getCount` obtient le nombre de tableaux dans la collection.</span><span class="sxs-lookup"><span data-stu-id="44bef-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="44bef-165">`getCount` renvoie un `ClientResult<number>` , ce qui signifie que la propriété dans le retour est un `value` [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) nombre.</span><span class="sxs-lookup"><span data-stu-id="44bef-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) is a number.</span></span> <span data-ttu-id="44bef-166">Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé.</span><span class="sxs-lookup"><span data-stu-id="44bef-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="44bef-167">À l’instar du chargement d’une propriété, la valeur `value` est une valeur « vide » locale jusqu’à cet appel`sync`.</span><span class="sxs-lookup"><span data-stu-id="44bef-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="44bef-168">Le script suivant fournit le nombre total de tableaux dans le classeur et enregistre ce nombre sur la console.</span><span class="sxs-lookup"><span data-stu-id="44bef-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

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

## <a name="converting-async-scripts-to-the-current-model"></a><span data-ttu-id="44bef-169">Conversion de scripts async en modèle actuel</span><span class="sxs-lookup"><span data-stu-id="44bef-169">Converting async scripts to the current model</span></span>

<span data-ttu-id="44bef-170">Le modèle API actuel n’utilise `load` pas , ou un `sync` `RequestContext` .</span><span class="sxs-lookup"><span data-stu-id="44bef-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="44bef-171">Cela facilite l’écriture et la maintenance des scripts.</span><span class="sxs-lookup"><span data-stu-id="44bef-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="44bef-172">Votre meilleure ressource pour la conversion d’anciens scripts est [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="44bef-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="44bef-173">Là, vous pouvez demander de l’aide à la communauté pour des scénarios spécifiques.</span><span class="sxs-lookup"><span data-stu-id="44bef-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="44bef-174">Les instructions suivantes doivent vous aider à décrire les étapes générales que vous devez suivre.</span><span class="sxs-lookup"><span data-stu-id="44bef-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="44bef-175">Créez un script et copiez-y l’ancien code async.</span><span class="sxs-lookup"><span data-stu-id="44bef-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="44bef-176">Assurez-vous de ne pas inclure l’ancienne `main` signature de méthode, en utilisant la signature `function main(workbook: ExcelScript.Workbook)` actuelle à la place.</span><span class="sxs-lookup"><span data-stu-id="44bef-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="44bef-177">Supprimez tous les `load` `sync` appels.</span><span class="sxs-lookup"><span data-stu-id="44bef-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="44bef-178">Elles ne sont plus nécessaires.</span><span class="sxs-lookup"><span data-stu-id="44bef-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="44bef-179">Toutes les propriétés ont été supprimées.</span><span class="sxs-lookup"><span data-stu-id="44bef-179">All properties have been removed.</span></span> <span data-ttu-id="44bef-180">Vous accédez maintenant à ces objets par le biais de méthodes, vous devrez donc basculer ces références de propriétés vers des appels `get` `set` de méthode.</span><span class="sxs-lookup"><span data-stu-id="44bef-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="44bef-181">Par exemple, au lieu de définir la couleur de remplissage d’une cellule par le biais de l’accès aux propriétés comme ceci : , vous allez maintenant utiliser des méthodes `mySheet.getRange("A2:C2").format.fill.color = "blue";` comme celle-ci : `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="44bef-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="44bef-182">Les classes de collection ont été remplacées par des tableaux.</span><span class="sxs-lookup"><span data-stu-id="44bef-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="44bef-183">Les méthodes et les méthodes de ces classes de collection ont été déplacées vers l’objet propriétaire de la collection. Vos références doivent donc être mises `add` `get` à jour en conséquence.</span><span class="sxs-lookup"><span data-stu-id="44bef-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="44bef-184">Par exemple, pour obtenir un graphique nommé « MyChart » à partir de la première feuille de calcul du manuel, utilisez le code suivant `workbook.getWorksheets()[0].getChart("MyChart");` :</span><span class="sxs-lookup"><span data-stu-id="44bef-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="44bef-185">Notez que `[0]` pour accéder à la première valeur de la valeur `Worksheet[]` renvoyée par `getWorksheets()` .</span><span class="sxs-lookup"><span data-stu-id="44bef-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="44bef-186">Certaines méthodes ont été renommées pour plus de clarté et ajoutées par souci de commodité.</span><span class="sxs-lookup"><span data-stu-id="44bef-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="44bef-187">Pour plus [d’informations, consultez la référence de l’API Office Scripts.](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="44bef-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="44bef-188">Documentation de référence de l’API async des scripts Office</span><span class="sxs-lookup"><span data-stu-id="44bef-188">Office Scripts async API reference documentation</span></span>

<span data-ttu-id="44bef-189">Les API async sont équivalentes à celles utilisées dans les add-ins Office. La documentation de référence se trouve dans la section Excel de la référence de l’API JavaScript pour les [add-ins Office.](/javascript/api/excel?view=excel-js-online&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="44bef-189">The async APIs are equivalent to those used in Office Add-ins. The reference documentation is found in [the Excel section of the Office Add-ins JavaScript API reference](/javascript/api/excel?view=excel-js-online&preserve-view=true).</span></span>
