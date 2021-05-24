---
title: Utiliser les appels externes de récupération à Office Scripts
description: Découvrez comment effectuer des appels d’API externes dans Office scripts.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: df8814cbab16969a1140aecfe526fd68e609d43c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545751"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="b8005-103">Utiliser les appels externes de récupération à Office Scripts</span><span class="sxs-lookup"><span data-stu-id="b8005-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="b8005-104">Ce script obtient des informations de base sur les référentiels GitHub d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b8005-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="b8005-105">Il montre comment utiliser `fetch` dans un scénario simple.</span><span class="sxs-lookup"><span data-stu-id="b8005-105">It shows how to use `fetch` in a simple scenario.</span></span> <span data-ttu-id="b8005-106">Pour plus d’informations sur l’utilisation ou d’autres appels externes, lisez la prise en charge des appels d’API externes `fetch` [dans Office Scripts](../../develop/external-calls.md)</span><span class="sxs-lookup"><span data-stu-id="b8005-106">For more information about using `fetch` or other external calls, read [External API call support in Office Scripts](../../develop/external-calls.md)</span></span>

<span data-ttu-id="b8005-107">Vous pouvez en savoir plus sur les API GItHub utilisées dans la référence GitHub [API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)</span><span class="sxs-lookup"><span data-stu-id="b8005-107">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="b8005-108">Vous pouvez également voir la sortie d’appel d’API brute en visitant un navigateur web (n’oubliez pas de remplacer l’espace réservé {USERNAME} par votre `https://api.github.com/users/{USERNAME}/repos` ID GitHub).</span><span class="sxs-lookup"><span data-stu-id="b8005-108">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your GitHub ID).</span></span>

![Obtenir un exemple d’informations sur les référentiels](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="b8005-110">Exemple de code : obtenir des informations de base sur les référentiels GitHub utilisateur</span><span class="sxs-lookup"><span data-stu-id="b8005-110">Sample code: Get basic information about user's GitHub repositories</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Call the GitHub REST API.
  // Replace the {USERNAME} placeholder with your GitHub username.
  const response = await fetch('https://api.github.com/users/{USERNAME}/repos');
  const repos: Repository[] = await response.json();
  
  // Create an array to hold the returned values.
  const rows: (string | boolean | number)[][] = [];

  // Convert each repository block into a row.
  for (let repo of repos){ 
    rows.push([repo.id, repo.name, repo.license?.name, repo.license?.url])
  }

  // Add the data to the current worksheet, starting at "A2".
  const sheet = workbook.getActiveWorksheet();
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
}

// An interface matching the returned JSON for a GitHub repository.
interface Repository {
  name: string,
  id: string,
  license?: License 
}

// An interface matching the returned JSON for a GitHub repo license.
interface License {
  name: string,
  url: string
}
```

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="b8005-111">Vidéo de formation : Comment effectuer des appels d’API externes</span><span class="sxs-lookup"><span data-stu-id="b8005-111">Training video: How to make external API calls</span></span>

<span data-ttu-id="b8005-112">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/fulP29J418E).</span><span class="sxs-lookup"><span data-stu-id="b8005-112">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
