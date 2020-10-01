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
# <a name="external-api-call-support-in-office-scripts"></a>Prise en charge des appels d’API externes dans Scripts Office

La plateforme de scripts Office ne prend pas en charge les appels vers des [API externes](https://developer.mozilla.org/docs/Web/API). Toutefois, ces appels peuvent être exécutés dans les bonnes circonstances. Les appels externes ne peuvent être effectués qu’à travers le client Excel, et non par le biais de la mise à l’arrêt automatique [dans des circonstances normales](#external-calls-from-power-automate).

Les auteurs de script ne devraient pas attendre un comportement cohérent lors de l’utilisation d’API externes pendant la phase d’aperçu de la plateforme. Cela est dû à la façon dont le runtime JavaScript gère l’interaction avec le classeur. Le script peut se terminer avant la fin de l’appel de l’API (ou sa `Promise` résolution est entièrement résolue). En tant que telles, ne reposez pas sur les API externes pour les scénarios de scripts critiques.

> [!CAUTION]
> Les appels externes peuvent entraîner l’exposition des données sensibles à des points de terminaison indésirables. Votre administrateur peut établir une protection de pare-feu contre ces appels.

## <a name="definition-files-for-external-apis"></a>Fichiers de définition pour les API externes

Les fichiers de définition des API externes ne sont pas inclus dans les scripts Office. L’utilisation de ces API génère des erreurs de compilation pour les définitions manquantes. Les API continuent à s’exécuter (même si elles sont exécutées via le client Excel), comme illustré dans le script suivant :

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

## <a name="external-calls-from-power-automate"></a>Appels externes de Power Automated

Les appels de l’API externe échouent lorsqu’un script est exécuté avec Power Automated. Il s’agit d’une différence de comportement entre l’exécution d’un script via le client Excel et l’automatisation de l’alimentation. Veillez à vérifier les scripts de ces références avant de les générer dans un flux.

> [!WARNING]
> Échec des appels externes le [connecteur Excel Online](/connectors/excelonlinebusiness) de Power automate est là pour vous aider à respecter les stratégies de protection contre la perte de données existantes. Toutefois, les scripts exécutés à l’aide de l’automate d’alimentation sont effectués de façon extérieure à votre organisation et en dehors des pare-feu de votre organisation. Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation des scripts Office. Votre administrateur peut désactiver le connecteur Excel Online dans Power automate ou désactiver les scripts Office pour Excel sur le Web via les contrôles de l' [administrateur des scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Voir aussi

- [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)