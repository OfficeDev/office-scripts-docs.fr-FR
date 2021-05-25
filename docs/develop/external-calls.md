---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Prise en charge et conseils pour effectuer des appels d’API externes dans Office Script.
ms.date: 05/21/2021
localization_priority: Normal
ms.openlocfilehash: 5d768b53112473c1774f8fe8257b197ffead4a63
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631642"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="f9505-103">Prise en charge des appels d’API externes dans Scripts Office</span><span class="sxs-lookup"><span data-stu-id="f9505-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="f9505-104">Les scripts prendre en charge les appels à des services externes.</span><span class="sxs-lookup"><span data-stu-id="f9505-104">Scripts support calls to external services.</span></span> <span data-ttu-id="f9505-105">Utilisez ces services pour fournir des données et d’autres informations à votre workbook.</span><span class="sxs-lookup"><span data-stu-id="f9505-105">Use these services to supply data and other information to your workbook.</span></span>

> [!CAUTION]
> <span data-ttu-id="f9505-106">Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables.</span><span class="sxs-lookup"><span data-stu-id="f9505-106">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="f9505-107">Votre administrateur peut établir une protection pare-feu contre ces appels.</span><span class="sxs-lookup"><span data-stu-id="f9505-107">Your admin can establish firewall protection against such calls.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f9505-108">Les appels aux API externes peuvent uniquement être effectués via l’application Excel, et non Power Automate [dans des circonstances normales.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="f9505-108">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

## <a name="configure-your-script-for-external-calls"></a><span data-ttu-id="f9505-109">Configurer votre script pour les appels externes</span><span class="sxs-lookup"><span data-stu-id="f9505-109">Configure your script for external calls</span></span>

<span data-ttu-id="f9505-110">Les appels externes [sont asynchrones](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) et nécessitent que votre script soit marqué comme `async` .</span><span class="sxs-lookup"><span data-stu-id="f9505-110">External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`.</span></span> <span data-ttu-id="f9505-111">Ajoutez `async` le préfixe à votre fonction et lui `main` renvoyez un , comme illustré ici `Promise` :</span><span class="sxs-lookup"><span data-stu-id="f9505-111">Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> <span data-ttu-id="f9505-112">Les scripts qui retournent d’autres informations peuvent `Promise` renvoyer un type de ce type.</span><span class="sxs-lookup"><span data-stu-id="f9505-112">Scripts that return other information can return a `Promise` of that type.</span></span> <span data-ttu-id="f9505-113">Par exemple, si votre script doit renvoyer un `Employee` objet, la signature de retour sera `: Promise <Employee>`</span><span class="sxs-lookup"><span data-stu-id="f9505-113">For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`</span></span>

<span data-ttu-id="f9505-114">Vous devez découvrir les interfaces du service externe pour appeler ce service.</span><span class="sxs-lookup"><span data-stu-id="f9505-114">You'll need to learn the external service's interfaces to make calls to that service.</span></span> <span data-ttu-id="f9505-115">Si vous utilisez ou des API REST, vous devez déterminer la `fetch` structure JSON des données renvoyées. [](https://wikipedia.org/wiki/Representational_state_transfer)</span><span class="sxs-lookup"><span data-stu-id="f9505-115">If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data.</span></span> <span data-ttu-id="f9505-116">Pour l’entrée et la sortie de votre script, envisagez d’effectuer une correspondance avec `interface` les structures JSON nécessaires.</span><span class="sxs-lookup"><span data-stu-id="f9505-116">For both input to and output from your script, consider making an `interface` to match the needed JSON structures.</span></span> <span data-ttu-id="f9505-117">Cela permet au script d’améliorer la sécurité des types.</span><span class="sxs-lookup"><span data-stu-id="f9505-117">This gives the script more type safety.</span></span> <span data-ttu-id="f9505-118">Vous pouvez en voir un exemple dans [l’utilisation de la récupération à partir Office scripts](../resources/samples/external-fetch-calls.md).</span><span class="sxs-lookup"><span data-stu-id="f9505-118">You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).</span></span>

### <a name="limitations-with-external-calls-from-office-scripts"></a><span data-ttu-id="f9505-119">Limitations avec les appels externes de Office Scripts</span><span class="sxs-lookup"><span data-stu-id="f9505-119">Limitations with external calls from Office Scripts</span></span>

* <span data-ttu-id="f9505-120">Il n’existe aucun moyen de se connecter ou d’utiliser le type de flux d’authentification OAuth2.</span><span class="sxs-lookup"><span data-stu-id="f9505-120">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="f9505-121">Toutes les clés et informations d’identification doivent être codées en dur (ou lues à partir d’une autre source).</span><span class="sxs-lookup"><span data-stu-id="f9505-121">All keys and credentials have to be hardcoded (or read from another source).</span></span>
* <span data-ttu-id="f9505-122">Il n’existe aucune infrastructure pour stocker les informations d’identification et les clés d’API.</span><span class="sxs-lookup"><span data-stu-id="f9505-122">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="f9505-123">Il devra être géré par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f9505-123">This will have to be managed by the user.</span></span>
* <span data-ttu-id="f9505-124">Les cookies de document `localStorage` et les objets ne sont pas pris en `sessionStorage` charge.</span><span class="sxs-lookup"><span data-stu-id="f9505-124">Document cookies, `localStorage`, and `sessionStorage` objects are not supported.</span></span>
* <span data-ttu-id="f9505-125">Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables ou des données externes à mettre dans des workbooks internes.</span><span class="sxs-lookup"><span data-stu-id="f9505-125">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="f9505-126">Votre administrateur peut établir une protection pare-feu contre ces appels.</span><span class="sxs-lookup"><span data-stu-id="f9505-126">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="f9505-127">Veillez à vérifier les stratégies locales avant de vous appuyer sur des appels externes.</span><span class="sxs-lookup"><span data-stu-id="f9505-127">Be sure to check with local policies prior to relying on external calls.</span></span>
* <span data-ttu-id="f9505-128">Veillez à vérifier la quantité de débit de données avant de prendre une dépendance.</span><span class="sxs-lookup"><span data-stu-id="f9505-128">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="f9505-129">Par exemple, il est possible que le fait d’retirer l’intégralité du jeu de données externe ne soit pas la meilleure option et que la pagination soit utilisée pour obtenir des données par blocs.</span><span class="sxs-lookup"><span data-stu-id="f9505-129">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="retrieve-information-with-fetch"></a><span data-ttu-id="f9505-130">Récupérer des informations avec `fetch`</span><span class="sxs-lookup"><span data-stu-id="f9505-130">Retrieve information with `fetch`</span></span>

<span data-ttu-id="f9505-131">[L’API de](https://developer.mozilla.org/docs/Web/API/Fetch_API) récupération récupère des informations à partir de services externes.</span><span class="sxs-lookup"><span data-stu-id="f9505-131">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="f9505-132">Il s’agit `async` d’une API, vous devez donc ajuster la `main` signature de votre script.</span><span class="sxs-lookup"><span data-stu-id="f9505-132">It is an `async` API, so you need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="f9505-133">Make the `main` function and have it return a `async` `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="f9505-133">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="f9505-134">Vous devez également être sûr de `await` `fetch` l’appel et de la `json` récupération.</span><span class="sxs-lookup"><span data-stu-id="f9505-134">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="f9505-135">Cela garantit que ces opérations sont terminées avant la fin du script.</span><span class="sxs-lookup"><span data-stu-id="f9505-135">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="f9505-136">Toutes les données JSON récupérées par `fetch` doivent correspondre à une interface définie dans le script.</span><span class="sxs-lookup"><span data-stu-id="f9505-136">Any JSON data retrieved by `fetch` must match an interface defined in the script.</span></span> <span data-ttu-id="f9505-137">La valeur renvoyée doit être affectée à un type spécifique, [car les scripts Office ne le supportent `any` pas.](typescript-restrictions.md#no-any-type-in-office-scripts)</span><span class="sxs-lookup"><span data-stu-id="f9505-137">The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts).</span></span> <span data-ttu-id="f9505-138">Vous devez consulter la documentation de votre service pour voir les noms et les types des propriétés renvoyées.</span><span class="sxs-lookup"><span data-stu-id="f9505-138">You should refer to the documentation for your service to see what the names and types of the returned properties are.</span></span> <span data-ttu-id="f9505-139">Ensuite, ajoutez l’interface ou les interfaces correspondantes à votre script.</span><span class="sxs-lookup"><span data-stu-id="f9505-139">Then, add the matching interface or interfaces to your script.</span></span>

<span data-ttu-id="f9505-140">Le script suivant utilise `fetch` pour récupérer les données JSON du serveur de test dans l’URL donnée.</span><span class="sxs-lookup"><span data-stu-id="f9505-140">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span> <span data-ttu-id="f9505-141">Notez `JSONData` l’interface pour stocker les données en tant que type correspondant.</span><span class="sxs-lookup"><span data-stu-id="f9505-141">Note the `JSONData` interface to store the data as a matching type.</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a><span data-ttu-id="f9505-142">Autres `fetch` exemples</span><span class="sxs-lookup"><span data-stu-id="f9505-142">Other `fetch` samples</span></span>

* <span data-ttu-id="f9505-143">L’exemple Utiliser des appels de récupération externe [dans Office Scripts](../resources/samples/external-fetch-calls.md) montre comment obtenir des informations de base sur les référentiels GitHub d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f9505-143">The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.</span></span>
* <span data-ttu-id="f9505-144">L’exemple de scénario [Office Scripts](../resources/scenarios/noaa-data-fetch.md) : Graph données au niveau de l’eau de la NOAA illustre la commande de récupération utilisée pour extraire des enregistrements de la base de données Archives et courants de l’administration nationale.</span><span class="sxs-lookup"><span data-stu-id="f9505-144">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="f9505-145">Appels externes de Power Automate</span><span class="sxs-lookup"><span data-stu-id="f9505-145">External calls from Power Automate</span></span>

<span data-ttu-id="f9505-146">Tout appel d’API externe échoue lorsqu’un script est exécuté avec Power Automate.</span><span class="sxs-lookup"><span data-stu-id="f9505-146">Any external API call fails when a script is run with Power Automate.</span></span> <span data-ttu-id="f9505-147">Il s’agit d’une différence de comportement entre l’exécution d’un script via l Excel’application Power Automate.</span><span class="sxs-lookup"><span data-stu-id="f9505-147">This is a behavioral difference between running a script through the Excel application and through Power Automate.</span></span> <span data-ttu-id="f9505-148">Veillez à vérifier si vos scripts sont de telles références avant de les créer dans un flux.</span><span class="sxs-lookup"><span data-stu-id="f9505-148">Be sure to check your scripts for such references before building them into a flow.</span></span>

<span data-ttu-id="f9505-149">Vous devez utiliser HTTP avec [Azure AD](/connectors/webcontents/) ou d’autres actions équivalentes pour tirer des données ou les pousser vers un service externe.</span><span class="sxs-lookup"><span data-stu-id="f9505-149">You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.</span></span>

> [!WARNING]
> <span data-ttu-id="f9505-150">Les appels externes effectués via le connecteur Power Automate [Excel Online](/connectors/excelonlinebusiness) échouent pour aider à respecter les stratégies de protection contre la perte de données existantes.</span><span class="sxs-lookup"><span data-stu-id="f9505-150">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="f9505-151">Toutefois, les scripts exécutés par Power Automate sont effectués en dehors de votre organisation et en dehors des pare-feu de votre organisation.</span><span class="sxs-lookup"><span data-stu-id="f9505-151">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="f9505-152">Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation Office scripts.</span><span class="sxs-lookup"><span data-stu-id="f9505-152">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="f9505-153">Votre administrateur peut désactiver le connecteur Excel Online dans Power Automate ou désactiver les scripts Office pour Excel sur le Web via les contrôles d’administrateur [Office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="f9505-153">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="f9505-154">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f9505-154">See also</span></span>

* [<span data-ttu-id="f9505-155">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="f9505-155">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
* [<span data-ttu-id="f9505-156">Utiliser les appels externes de récupération dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="f9505-156">Use external fetch calls in Office Scripts</span></span>](../resources/samples/external-fetch-calls.md)
* [<span data-ttu-id="f9505-157">Office Exemple de scénario de scripts : Graph données de niveau d’eau à partir de la NOAA</span><span class="sxs-lookup"><span data-stu-id="f9505-157">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
