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
# <a name="external-api-calls-from-office-scripts"></a>Appels d'API externes à partir de Scripts Office

Les scripts Office permettent une [prise en charge limitée des appels d'API externes.](../../develop/external-calls.md)

> [!IMPORTANT]
>
> * Il n'existe aucun moyen de se connecter ou d'utiliser le type de flux d'authentification OAuth2. Toutes les clés et informations d'identification doivent être codées en dur (ou lues à partir d'une autre source).
> * Il n'existe aucune infrastructure pour stocker les informations d'identification et les clés d'API. Il devra être géré par l'utilisateur.
> * Les appels externes peuvent entraîner l'exposition de données sensibles à des points de terminaison indésirables ou des données externes à entrer dans des workbooks internes. Votre administrateur peut établir une protection pare-feu contre ces appels. Veillez à vérifier les stratégies locales avant de vous appuyer sur des appels externes.
> * Si un script utilise un appel d'API, il ne fonctionne pas dans un scénario Power Automate. Vous devez utiliser l'action HTTP de Power Automate ou des actions équivalentes pour tirer des données ou les pousser vers un service externe.
> * Un appel d'API externe implique une syntaxe d'API asynchrone et nécessite une connaissance légèrement avancée du fonctionnement de la communication asynchrone.
> * Veillez à vérifier la quantité de débit de données avant de prendre une dépendance. Par exemple, il est possible que le fait d'retirer l'intégralité du jeu de données externe ne soit pas la meilleure option et que la pagination soit utilisée pour obtenir des données par blocs.

## <a name="useful-knowledge-and-resources"></a>Connaissances et ressources utiles

* [API REST](https://en.wikipedia.org/wiki/Representational_state_transfer): la façon la plus probable d'utiliser l'appel d'API.
* [ `async` : comprendre comment cela fonctionne. `await` ](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await)
* [`fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API/Using_Fetch): comprendre comment cela fonctionne.

## <a name="steps"></a>Étapes

1. Marquez `main` votre fonction comme une fonction asynchrone en ajoutant un `async` préfixe. Par exemple, `async function main(workbook: ExcelScript.Workbook)`.
1. Quel type d'appel API faites-vous ? `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? Pour plus d'informations, reportez-vous au matériel de l'API REST.
1. Obtenez le point de terminaison de l'API de service, les exigences d'authentification, les en-têtes, etc.
1. Définissez l'entrée ou la sortie `interface` pour faciliter la vérification de l'achèvement du code et du temps de développement. Pour [plus d'informations,](#training-video-how-to-make-external-api-calls) voir la vidéo.
1. Code, test, optimiser. Vous pouvez créer une fonction pour votre routine d'appel d'API pour la rendre réutilisable à partir d'autres parties de votre script ou pour la réutiliser dans un autre script (copier-coller devient beaucoup plus facile de cette façon).

## <a name="scenario"></a>Scénario

Ce script obtient des informations de base sur les référentiels GitHub de l'utilisateur.

## <a name="resources-used-in-the-sample"></a>Ressources utilisées dans l'exemple

1. [Obtenir la référence de l'API Github des référentiels.](https://docs.github.com/rest/reference/repos#list-repositories-for-a-user)
1. Sortie d'appel d'API : go to a web browser or any HTTP interface and type in `https://api.github.com/users/{USERNAME}/repos` , replacing the {USERNAME} placeholder with your Github ID.
1. Informations récupérées : repo.name, repo.size, repo.owner.id, repo.license?. name

## <a name="sample-code-get-basic-information-about-users-github-repositories"></a>Exemple de code : obtenir des informations de base sur les référentiels GitHub de l'utilisateur

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

## <a name="training-video-how-to-make-external-api-calls"></a>Vidéo de formation : Comment effectuer des appels d'API externes

[![Regarder une vidéo sur la façon d'effectuer des appels d'API externes](../../images/api-vid.png)](https://youtu.be/fulP29J418E "Vidéo sur la façon d'effectuer des appels d'API externes")
