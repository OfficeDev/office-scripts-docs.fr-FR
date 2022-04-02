---
title: Utiliser les appels externes de récupération dans les scripts Office
description: Découvrez comment effectuer des appels d’API externes dans Office scripts.
ms.date: 05/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: feff9d49f9f50f14fd83b1864568df8dab02d417
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585526"
---
# <a name="use-external-fetch-calls-in-office-scripts"></a>Utiliser les appels externes de récupération dans les scripts Office

Ce script obtient des informations de base sur les référentiels GitHub d’un utilisateur. Il montre comment utiliser dans `fetch` un scénario simple. Pour plus d’informations sur l’utilisation `fetch` ou d’autres appels externes, lisez la prise en charge des appels d’API externes [dans Office Scripts](../../develop/external-calls.md)

Vous pouvez en savoir plus sur les API GItHub utilisées dans la référence GitHub [API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) Vous pouvez également voir la sortie d’appel d’API `https://api.github.com/users/{USERNAME}/repos` brute en vous rendant dans un navigateur web (n’oubliez pas de remplacer l’espace réservé {USERNAME} par votre ID GitHub).

![Obtenir un exemple d’informations sur les référentiels](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Exemple de code : obtenir des informations de base sur les référentiels GitHub utilisateur

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

## <a name="training-video-how-to-make-external-api-calls"></a>Vidéo de formation : Comment effectuer des appels d’API externes

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/fulP29J418E).
