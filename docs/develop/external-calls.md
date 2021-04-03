---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Prise en charge et conseils pour effectuer des appels d’API externes dans un script Office.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: 74b8750f609370370759ca4a4a1daa998363ac2e
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570310"
---
# <a name="external-api-call-support-in-office-scripts"></a>Prise en charge des appels d’API externes dans Scripts Office

Les auteurs de scripts ne doivent pas s’attendre à un comportement cohérent lors de l’utilisation [d’API](https://developer.mozilla.org/docs/Web/API) externes lors de la phase de prévisualisation de la plateforme. En tant que tel, ne comptez pas sur les API externes pour les scénarios de script critiques.

Les appels aux API externes peuvent uniquement être effectués via l’application Excel, et non via Power Automate [dans des circonstances normales.](#external-calls-from-power-automate)

> [!CAUTION]
> Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables. Votre administrateur peut établir une protection pare-feu contre ces appels.

## <a name="working-with-fetch"></a>Travailler avec `fetch`

[L’API de](https://developer.mozilla.org/docs/Web/API/Fetch_API) récupération récupère des informations à partir de services externes. Il s’agit `async` d’une API, vous devrez donc ajuster la `main` signature de votre script. Make the `main` function and have it return a `async` `Promise<void>` . Vous devez également être sûr de `await` `fetch` l’appel et de la `json` récupération. Cela garantit que ces opérations sont terminées avant la fin du script.

Le script suivant utilise `fetch` pour récupérer les données JSON du serveur de test dans l’URL donnée.

```TypeScript
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

L’exemple de scénario Office Scripts : les données de niveau d’eau graphe de [la NOAA](../resources/scenarios/noaa-data-fetch.md) illustrent la commande de récupération utilisée pour extraire des enregistrements de la base de données Archives et courants de l’Administration nationale.

## <a name="external-calls-from-power-automate"></a>Appels externes de Power Automate

Tous les appels d’API externes échouent lorsqu’un script est exécuté avec Power Automate. Il s’agit d’une différence comportementale entre l’exécution d’un script via le client Excel et via Power Automate. Veillez à vérifier si vos scripts sont de telles références avant de les créer dans un flux.

> [!WARNING]
> Les appels externes effectués via le connecteur Power Automate [Excel Online](/connectors/excelonlinebusiness) échouent afin d’aider à respecter les stratégies de protection contre la perte de données existantes. Toutefois, les scripts exécutés via Power Automate le sont en dehors de votre organisation et en dehors des pare-feu de votre organisation. Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation des scripts Office. Votre administrateur peut désactiver le connecteur Excel Online dans Power Automate ou désactiver les scripts Office pour Excel sur le web via les contrôles d’administrateur des [scripts Office.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Voir aussi

- [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
- [Exemple de scénario de scripts Office : graphique des données de niveau d’eau de NOAA](../resources/scenarios/noaa-data-fetch.md)
