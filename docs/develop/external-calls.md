---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Prise en charge et conseils pour effectuer des appels d’API externes dans un script Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 1091031bc2e12f3e1e79b177c69874ee4ce61dd8
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784143"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="3445f-103">Prise en charge des appels d’API externes dans Scripts Office</span><span class="sxs-lookup"><span data-stu-id="3445f-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="3445f-104">Les auteurs de scripts ne doivent pas s’attendre à un comportement cohérent lors de l’utilisation [d’API](https://developer.mozilla.org/docs/Web/API) externes lors de la phase de prévisualisation de la plateforme.</span><span class="sxs-lookup"><span data-stu-id="3445f-104">Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase.</span></span> <span data-ttu-id="3445f-105">En tant que tel, ne comptez pas sur les API externes pour les scénarios de script critiques.</span><span class="sxs-lookup"><span data-stu-id="3445f-105">As such, do not rely on external APIs for critical script scenarios.</span></span>

<span data-ttu-id="3445f-106">Les appels aux API externes peuvent uniquement être effectués via l’application Excel, et non via Power Automate [dans des circonstances normales.](#external-calls-from-power-automate)</span><span class="sxs-lookup"><span data-stu-id="3445f-106">Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

> [!CAUTION]
> <span data-ttu-id="3445f-107">Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables.</span><span class="sxs-lookup"><span data-stu-id="3445f-107">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="3445f-108">Votre administrateur peut établir une protection pare-feu contre ces appels.</span><span class="sxs-lookup"><span data-stu-id="3445f-108">Your admin can establish firewall protection against such calls.</span></span>

## <a name="working-with-fetch"></a><span data-ttu-id="3445f-109">Travailler avec `fetch`</span><span class="sxs-lookup"><span data-stu-id="3445f-109">Working with `fetch`</span></span>

<span data-ttu-id="3445f-110">[L’API de](https://developer.mozilla.org/docs/Web/API/Fetch_API) récupération récupère des informations à partir de services externes.</span><span class="sxs-lookup"><span data-stu-id="3445f-110">The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services.</span></span> <span data-ttu-id="3445f-111">Il s’agit `async` d’une API, vous devrez donc ajuster la `main` signature de votre script.</span><span class="sxs-lookup"><span data-stu-id="3445f-111">It is an `async` API, so you will need to adjust the `main` signature of your script.</span></span> <span data-ttu-id="3445f-112">Make the `main` function and have it return a `async` `Promise<void>` .</span><span class="sxs-lookup"><span data-stu-id="3445f-112">Make the `main` function `async` and have it return a `Promise<void>`.</span></span> <span data-ttu-id="3445f-113">Vous devez également être sûr de `await` `fetch` l’appel et de la `json` récupération.</span><span class="sxs-lookup"><span data-stu-id="3445f-113">You should also be sure to `await` the `fetch` call and `json` retrieval.</span></span> <span data-ttu-id="3445f-114">Cela garantit que ces opérations sont terminées avant la fin du script.</span><span class="sxs-lookup"><span data-stu-id="3445f-114">This ensures those operations complete before the script ends.</span></span>

<span data-ttu-id="3445f-115">Le script suivant utilise `fetch` pour récupérer les données JSON du serveur de test dans l’URL donnée.</span><span class="sxs-lookup"><span data-stu-id="3445f-115">The following script uses `fetch` to retrieve JSON data from the test server in the given URL.</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* 
   * Retrieve JSON data from a test server.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

<span data-ttu-id="3445f-116">L’exemple de scénario Office Scripts : graphe des données de niveau d’eau de [la NOAA](../resources/scenarios/noaa-data-fetch.md) illustre la commande de récupération utilisée pour extraire des enregistrements de la base de données Archives et courants de l’Administration nationale.</span><span class="sxs-lookup"><span data-stu-id="3445f-116">The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.</span></span>

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="3445f-117">Appels externes à partir de Power Automate</span><span class="sxs-lookup"><span data-stu-id="3445f-117">External calls from Power Automate</span></span>

<span data-ttu-id="3445f-118">Tous les appels d’API externes échouent lorsqu’un script est exécuté avec Power Automate.</span><span class="sxs-lookup"><span data-stu-id="3445f-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="3445f-119">Il s’agit d’une différence comportementale entre l’exécution d’un script via le client Excel et via Power Automate.</span><span class="sxs-lookup"><span data-stu-id="3445f-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="3445f-120">Veillez à vérifier si vos scripts sont de telles références avant de les créer dans un flux.</span><span class="sxs-lookup"><span data-stu-id="3445f-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="3445f-121">Les appels externes effectués via le connecteur Power Automate [Excel Online](/connectors/excelonlinebusiness) échouent afin d’aider à respecter les stratégies de protection contre la perte de données existantes.</span><span class="sxs-lookup"><span data-stu-id="3445f-121">External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="3445f-122">Toutefois, les scripts exécutés via Power Automate le sont en dehors de votre organisation et en dehors des pare-feu de votre organisation.</span><span class="sxs-lookup"><span data-stu-id="3445f-122">However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="3445f-123">Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation des scripts Office.</span><span class="sxs-lookup"><span data-stu-id="3445f-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="3445f-124">Votre administrateur peut désactiver le connecteur Excel Online dans Power Automate ou désactiver les scripts Office pour Excel sur le web via les contrôles d’administrateur [des scripts Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="3445f-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="3445f-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3445f-125">See also</span></span>

- [<span data-ttu-id="3445f-126">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="3445f-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="3445f-127">Exemple de scénario de scripts Office : graphique des données de niveau d’eau de NOAA</span><span class="sxs-lookup"><span data-stu-id="3445f-127">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>](../resources/scenarios/noaa-data-fetch.md)
