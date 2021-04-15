---
title: Effectuer des appels d'API externes dans les scripts Office
description: Découvrez comment effectuer des appels d'API externes dans les scripts Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 0ed57ed3b97309dbb7ea196695dcc347e133b3cf
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754802"
---
# <a name="external-api-calls-from-office-scripts"></a><span data-ttu-id="b9a20-103">Appels d'API externes à partir de Scripts Office</span><span class="sxs-lookup"><span data-stu-id="b9a20-103">External API calls from Office Scripts</span></span>

<span data-ttu-id="b9a20-104">Les scripts Office permettent une [prise en charge limitée des appels d'API externes.](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="b9a20-104">Office Scripts allows [limited external API call support](../../develop/external-calls.md).</span></span>

> [!IMPORTANT]
>
> * <span data-ttu-id="b9a20-105">Il n'existe aucun moyen de se connecter ou d'utiliser le type de flux d'authentification OAuth2.</span><span class="sxs-lookup"><span data-stu-id="b9a20-105">There is no way to sign in or use OAuth2 type of authentication flows.</span></span> <span data-ttu-id="b9a20-106">Toutes les clés et informations d'identification doivent être codées en dur (ou lues à partir d'une autre source).</span><span class="sxs-lookup"><span data-stu-id="b9a20-106">All keys and credentials have to be hardcoded (or read from another source).</span></span>
> * <span data-ttu-id="b9a20-107">Il n'existe aucune infrastructure pour stocker les informations d'identification et les clés d'API.</span><span class="sxs-lookup"><span data-stu-id="b9a20-107">There is no infrastructure to store API credentials and keys.</span></span> <span data-ttu-id="b9a20-108">Il devra être géré par l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9a20-108">This will have to be managed by the user.</span></span>
> * <span data-ttu-id="b9a20-109">Les appels externes peuvent entraîner l'exposition de données sensibles à des points de terminaison indésirables ou des données externes à entrer dans des workbooks internes.</span><span class="sxs-lookup"><span data-stu-id="b9a20-109">External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks.</span></span> <span data-ttu-id="b9a20-110">Votre administrateur peut établir une protection pare-feu contre ces appels.</span><span class="sxs-lookup"><span data-stu-id="b9a20-110">Your admin can establish firewall protection against such calls.</span></span> <span data-ttu-id="b9a20-111">Veillez à vérifier les stratégies locales avant de vous appuyer sur des appels externes.</span><span class="sxs-lookup"><span data-stu-id="b9a20-111">Be sure to check with local policies prior to relying on external calls.</span></span>
> * <span data-ttu-id="b9a20-112">Si un script utilise un appel d'API, il ne fonctionne pas dans un scénario Power Automate.</span><span class="sxs-lookup"><span data-stu-id="b9a20-112">If a script uses an API call, it will not function in a Power Automate scenario.</span></span> <span data-ttu-id="b9a20-113">Vous devez utiliser l'action HTTP de Power Automate ou des actions équivalentes pour tirer des données ou les pousser vers un service externe.</span><span class="sxs-lookup"><span data-stu-id="b9a20-113">You'll have to use Power Automate's HTTP action or equivalent actions to pull data from or push it to an external service.</span></span>
> * <span data-ttu-id="b9a20-114">Un appel d'API externe implique une syntaxe d'API asynchrone et nécessite une connaissance légèrement avancée du fonctionnement de la communication asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b9a20-114">An external API call involves asynchronous API syntax and requires slightly advanced knowledge of the way async communication works.</span></span>
> * <span data-ttu-id="b9a20-115">Veillez à vérifier la quantité de débit de données avant de prendre une dépendance.</span><span class="sxs-lookup"><span data-stu-id="b9a20-115">Be sure to check the amount of data throughput prior to taking a dependency.</span></span> <span data-ttu-id="b9a20-116">Par exemple, il est possible que le fait d'retirer l'intégralité du jeu de données externe ne soit pas la meilleure option et que la pagination soit utilisée pour obtenir des données par blocs.</span><span class="sxs-lookup"><span data-stu-id="b9a20-116">For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.</span></span>

## <a name="useful-knowledge-and-resources"></a><span data-ttu-id="b9a20-117">Connaissances et ressources utiles</span><span class="sxs-lookup"><span data-stu-id="b9a20-117">Useful knowledge and resources</span></span>

* <span data-ttu-id="b9a20-118">[API REST](https://en.wikipedia.org/wiki/Representational_state_transfer): la façon la plus probable d'utiliser l'appel d'API.</span><span class="sxs-lookup"><span data-stu-id="b9a20-118">[REST API](https://en.wikipedia.org/wiki/Representational_state_transfer): Most likely way you'll use the API call.</span></span>
* <span data-ttu-id="b9a20-119">[ `async` : comprendre comment cela fonctionne. `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)</span><span class="sxs-lookup"><span data-stu-id="b9a20-119">[`async` `await`](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await): Understand how this works.</span></span>
* <span data-ttu-id="b9a20-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): comprendre comment cela fonctionne.</span><span class="sxs-lookup"><span data-stu-id="b9a20-120">[`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): Understand how this works.</span></span>

## <a name="steps"></a><span data-ttu-id="b9a20-121">Étapes</span><span class="sxs-lookup"><span data-stu-id="b9a20-121">Steps</span></span>

1. <span data-ttu-id="b9a20-122">Marquez `main` votre fonction comme une fonction asynchrone en ajoutant un `async` préfixe.</span><span class="sxs-lookup"><span data-stu-id="b9a20-122">Mark your `main` function as an asynchronous function by adding `async` prefix.</span></span> <span data-ttu-id="b9a20-123">Par exemple, `async function main(workbook: ExcelScript.Workbook)`.</span><span class="sxs-lookup"><span data-stu-id="b9a20-123">For example, `async function main(workbook: ExcelScript.Workbook)`.</span></span>
1. <span data-ttu-id="b9a20-124">Quel type d'appel API faites-vous ?</span><span class="sxs-lookup"><span data-stu-id="b9a20-124">Which type of API call are you making?</span></span> <span data-ttu-id="b9a20-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span><span class="sxs-lookup"><span data-stu-id="b9a20-125">`GET`, `POST`, `PUT`, `DELETE`, `PATCH`?</span></span> <span data-ttu-id="b9a20-126">Pour plus d'informations, reportez-vous au matériel de l'API REST.</span><span class="sxs-lookup"><span data-stu-id="b9a20-126">Refer to REST API material for details.</span></span>
1. <span data-ttu-id="b9a20-127">Obtenez le point de terminaison de l'API de service, les exigences d'authentification, les en-têtes, etc.</span><span class="sxs-lookup"><span data-stu-id="b9a20-127">Obtain the service API endpoint, authentication requirements, headers, etc.</span></span>
1. <span data-ttu-id="b9a20-128">Définissez l'entrée ou la sortie `interface` pour faciliter la vérification de l'achèvement du code et du temps de développement.</span><span class="sxs-lookup"><span data-stu-id="b9a20-128">Define the input or output `interface` to help with code completion and development time verification.</span></span> <span data-ttu-id="b9a20-129">Pour [plus d'informations,](#training-video-how-to-make-external-api-calls) voir la vidéo.</span><span class="sxs-lookup"><span data-stu-id="b9a20-129">See [video](#training-video-how-to-make-external-api-calls) for details.</span></span>
1. <span data-ttu-id="b9a20-130">Code, test, optimiser.</span><span class="sxs-lookup"><span data-stu-id="b9a20-130">Code, test, optimize.</span></span> <span data-ttu-id="b9a20-131">Vous pouvez créer une fonction pour votre routine d'appel d'API pour la rendre réutilisable à partir d'autres parties de votre script ou pour la réutiliser dans un autre script (copier-coller devient beaucoup plus facile de cette façon).</span><span class="sxs-lookup"><span data-stu-id="b9a20-131">You can create a function for your API call routine to make it reusable from other parts of your script or for reuse in a different script (copy-paste becomes much easier this way).</span></span>

## <a name="scenario"></a><span data-ttu-id="b9a20-132">Scénario</span><span class="sxs-lookup"><span data-stu-id="b9a20-132">Scenario</span></span>

<span data-ttu-id="b9a20-133">Ce script obtient des informations de base sur les référentiels GitHub de l'utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9a20-133">This script gets basic information about the user's GitHub repositories.</span></span>

## <a name="resources-used-in-the-sample"></a><span data-ttu-id="b9a20-134">Ressources utilisées dans l'exemple</span><span class="sxs-lookup"><span data-stu-id="b9a20-134">Resources used in the sample</span></span>

1. [<span data-ttu-id="b9a20-135">Obtenir la référence de l'API Github des référentiels.</span><span class="sxs-lookup"><span data-stu-id="b9a20-135">Get repositories Github API reference.</span></span>](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. <span data-ttu-id="b9a20-136">Sortie d'appel d'API : go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos` , replacing the {USERNAME} placeholder with your Github ID.</span><span class="sxs-lookup"><span data-stu-id="b9a20-136">API call output: Go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos`, replacing the {USERNAME} placeholder with your Github ID.</span></span>
1. <span data-ttu-id="b9a20-137">Informations récupérées : repo.name, repo.size, repo.owner.id, repo.license?. name</span><span class="sxs-lookup"><span data-stu-id="b9a20-137">Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name</span></span>

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="b9a20-138">Exemple de code : obtenir des informations de base sur les référentiels GitHub de l'utilisateur</span><span class="sxs-lookup"><span data-stu-id="b9a20-138">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {

  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  name: string,
  id: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="b9a20-139">Vidéo de formation : Comment effectuer des appels d'API externes</span><span class="sxs-lookup"><span data-stu-id="b9a20-139">Training video: How to make external API calls</span></span>

<span data-ttu-id="b9a20-140">[![Regarder une vidéo sur la façon d'effectuer des appels d'API externes](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Vidéo sur la façon d'effectuer des appels d'API externes")</span><span class="sxs-lookup"><span data-stu-id="b9a20-140">[![Watch video on how to make external API calls](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Video on how to make external API calls")</span></span>
