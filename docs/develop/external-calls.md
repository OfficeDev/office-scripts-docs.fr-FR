---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Prise en charge et conseils pour passer des appels d’API externes dans un script Office.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: fa77e606e2b3ab90144507660d71561b278e82e5
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319629"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="892b9-103">Prise en charge des appels d’API externes dans Scripts Office</span><span class="sxs-lookup"><span data-stu-id="892b9-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="892b9-104">La plateforme de scripts Office ne prend pas en charge les appels vers des [API externes](https://developer.mozilla.org/docs/Web/API).</span><span class="sxs-lookup"><span data-stu-id="892b9-104">The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API).</span></span> <span data-ttu-id="892b9-105">Toutefois, ces appels peuvent être exécutés dans les bonnes circonstances.</span><span class="sxs-lookup"><span data-stu-id="892b9-105">However, these calls can be run under the right circumstances.</span></span> <span data-ttu-id="892b9-106">Les appels externes ne peuvent être effectués qu’à travers le client Excel, et non par le biais de la mise à l’arrêt automatique [dans des circonstances normales](#external-calls-from-power-automate).</span><span class="sxs-lookup"><span data-stu-id="892b9-106">External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

<span data-ttu-id="892b9-107">Les auteurs de script ne devraient pas attendre un comportement cohérent lors de l’utilisation d’API externes pendant la phase d’aperçu de la plateforme.</span><span class="sxs-lookup"><span data-stu-id="892b9-107">Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.</span></span> <span data-ttu-id="892b9-108">Cela est dû à la façon dont le runtime JavaScript gère l’interaction avec le classeur.</span><span class="sxs-lookup"><span data-stu-id="892b9-108">This is due how the JavaScript runtime manages interacting with the workbook.</span></span> <span data-ttu-id="892b9-109">Le script peut se terminer avant la fin de l’appel de l’API (ou sa `Promise` résolution est entièrement résolue).</span><span class="sxs-lookup"><span data-stu-id="892b9-109">The script may end before the API call completes (or its `Promise` is fully resolved).</span></span> <span data-ttu-id="892b9-110">En tant que telles, ne reposez pas sur les API externes pour les scénarios de scripts critiques.</span><span class="sxs-lookup"><span data-stu-id="892b9-110">As such, do not rely on external APIs for critical script scenarios.</span></span>

> [!CAUTION]
> <span data-ttu-id="892b9-111">Les appels externes peuvent entraîner l’exposition des données sensibles à des points de terminaison indésirables.</span><span class="sxs-lookup"><span data-stu-id="892b9-111">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="892b9-112">Votre administrateur peut établir une protection de pare-feu contre ces appels.</span><span class="sxs-lookup"><span data-stu-id="892b9-112">Your admin can establish firewall protection against such calls.</span></span>

## <a name="definition-files-for-external-apis"></a><span data-ttu-id="892b9-113">Fichiers de définition pour les API externes</span><span class="sxs-lookup"><span data-stu-id="892b9-113">Definition files for external APIs</span></span>

<span data-ttu-id="892b9-114">Les fichiers de définition des API externes ne sont pas inclus dans les scripts Office.</span><span class="sxs-lookup"><span data-stu-id="892b9-114">The definition files for external APIs aren't included with Office Scripts.</span></span> <span data-ttu-id="892b9-115">L’utilisation de ces API génère des erreurs de compilation pour les définitions manquantes.</span><span class="sxs-lookup"><span data-stu-id="892b9-115">The use of such APIs generates compile-time errors for missing definitions.</span></span> <span data-ttu-id="892b9-116">Les API continuent à s’exécuter (même si elles sont exécutées via le client Excel), comme illustré dans le script suivant :</span><span class="sxs-lookup"><span data-stu-id="892b9-116">The APIs still run (though only when run through the Excel client), as shown in the following script:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="892b9-117">Appels externes de Power Automated</span><span class="sxs-lookup"><span data-stu-id="892b9-117">External calls from Power Automate</span></span>

<span data-ttu-id="892b9-118">Les appels de l’API externe échouent lorsqu’un script est exécuté avec Power Automated.</span><span class="sxs-lookup"><span data-stu-id="892b9-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="892b9-119">Il s’agit d’une différence de comportement entre l’exécution d’un script via le client Excel et l’automatisation de l’alimentation.</span><span class="sxs-lookup"><span data-stu-id="892b9-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="892b9-120">Veillez à vérifier les scripts de ces références avant de les générer dans un flux.</span><span class="sxs-lookup"><span data-stu-id="892b9-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="892b9-121">Échec des appels externes le [connecteur Excel Online](/connectors/excelonlinebusiness) de Power automate est là pour vous aider à respecter les stratégies de protection contre la perte de données existantes.</span><span class="sxs-lookup"><span data-stu-id="892b9-121">The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="892b9-122">Toutefois, les scripts exécutés à l’aide de l’automate d’alimentation sont effectués de façon extérieure à votre organisation et en dehors des pare-feu de votre organisation.</span><span class="sxs-lookup"><span data-stu-id="892b9-122">However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="892b9-123">Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation des scripts Office.</span><span class="sxs-lookup"><span data-stu-id="892b9-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="892b9-124">Votre administrateur peut désactiver le connecteur Excel Online dans Power automate ou désactiver les scripts Office pour Excel sur le Web via les contrôles de l' [administrateur des scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="892b9-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="892b9-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="892b9-125">See also</span></span>

- [<span data-ttu-id="892b9-126">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="892b9-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)