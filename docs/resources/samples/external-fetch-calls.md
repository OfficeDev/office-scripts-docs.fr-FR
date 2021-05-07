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
# <a name="use-external-fetch-calls-in-office-scripts"></a>Utiliser les appels externes de récupération à Office Scripts

Ce script obtient des informations de base sur les référentiels GitHub d’un utilisateur. Il montre comment utiliser `fetch` dans un scénario simple.

Vous pouvez en savoir plus sur les API GItHub utilisées dans la référence GitHub [API.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user) Vous pouvez également voir la sortie d’appel d’API brute en visitant un navigateur web (n’oubliez pas de remplacer l’espace réservé {USERNAME} par votre `https://api.github.com/users/{USERNAME}/repos` ID Github).

![Obtenir un exemple d’informations sur les référentiels](../../images/git.png)

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Exemple de code : obtenir des informations de base sur les référentiels GitHub utilisateur

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

## <a name="training-video-how-to-make-external-api-calls"></a>Vidéo de formation : Comment effectuer des appels d’API externes

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/fulP29J418E).