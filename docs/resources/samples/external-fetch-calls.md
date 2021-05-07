---
title: Utiliser les appels externes de récupération à Office Scripts
description: Découvrez comment effectuer des appels d’API externes dans Office scripts.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 721bfa39eea1e9973efc7fd13efa5bac734b76dd
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232521"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a><span data-ttu-id="be75b-103">Utiliser les appels externes de récupération à Office Scripts</span><span class="sxs-lookup"><span data-stu-id="be75b-103">Use external fetch calls in Office Scripts</span></span>

<span data-ttu-id="be75b-104">Ce script obtient des informations de base sur les référentiels GitHub d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be75b-104">This script gets basic information about a user's GitHub repositories.</span></span> <span data-ttu-id="be75b-105">Il montre comment utiliser `fetch` dans un scénario simple.</span><span class="sxs-lookup"><span data-stu-id="be75b-105">It shows how to use `fetch` in a simple scenario.</span></span>

<span data-ttu-id="be75b-106">Vous pouvez en savoir plus sur les API GItHub utilisées dans la référence GitHub [API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)</span><span class="sxs-lookup"><span data-stu-id="be75b-106">You can learn more about the GItHub APIs being used in the [GitHub API reference](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user).</span></span> <span data-ttu-id="be75b-107">Vous pouvez également voir la sortie d’appel d’API brute en visitant un navigateur web (n’oubliez pas de remplacer l’espace réservé {USERNAME} par votre `https://api.github.com/users/{USERNAME}/repos` ID Github).</span><span class="sxs-lookup"><span data-stu-id="be75b-107">You can also see the raw API call output by visiting `https://api.github.com/users/{USERNAME}/repos` in a web browser (be sure to replace the {USERNAME} placeholder with your Github ID).</span></span>

![Obtenir un exemple d’informations sur les référentiels](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a><span data-ttu-id="be75b-109">Exemple de code : obtenir des informations de base sur les référentiels GitHub utilisateur</span><span class="sxs-lookup"><span data-stu-id="be75b-109">Sample code: Get basic information about user's GitHub repositories</span></span>

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

## <a name="training-video-how-to-make-external-api-calls"></a><span data-ttu-id="be75b-110">Vidéo de formation : Comment effectuer des appels d’API externes</span><span class="sxs-lookup"><span data-stu-id="be75b-110">Training video: How to make external API calls</span></span>

<span data-ttu-id="be75b-111">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/fulP29J418E).</span><span class="sxs-lookup"><span data-stu-id="be75b-111">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/fulP29J418E).</span></span>
