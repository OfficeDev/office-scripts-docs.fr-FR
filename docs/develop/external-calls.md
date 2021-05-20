---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Support et conseils pour faire des appels API externes dans un script Office.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fd6ba0c57bf4cabb2d07421355cacff373f6706c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545081"
---
# <a name="external-api-call-support-in-office-scripts"></a>Prise en charge des appels d’API externes dans Scripts Office

Les auteurs de scripts ne devraient pas s’attendre à un comportement cohérent lors de [l’utilisation d’API](https://developer.mozilla.org/docs/Web/API) externes pendant la phase d’aperçu de la plate-forme. En tant que tel, ne vous fiez pas aux API externes pour les scénarios de script critiques.

Les appels vers des API externes ne peuvent être effectués que par le biais de l’application Excel, et non par le biais Power Automate [dans des circonstances normales.](#external-calls-from-power-automate)

> [!CAUTION]
> Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables. Votre administrateur peut établir une protection pare-feu contre de tels appels.

## <a name="configure-your-script-for-external-calls"></a>Configurez votre script pour des appels externes

Les appels externes [sont asynchrones et](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) exigent que votre script soit marqué comme `async` . Ajoutez le `async` préfixe à votre `main` fonction et qu’il retourne un `Promise` , comme indiqué ici:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Les scripts qui retournent d’autres informations peuvent renvoyer `Promise` un de ce type. Par exemple, si votre script doit retourner un `Employee` objet, la signature de retour serait `: Promise <Employee>`

Vous devrez apprendre les interfaces du service externe pour passer des appels vers ce service. Si vous utilisez ou `fetch` [REST API](https://wikipedia.org/wiki/Representational_state_transfer), vous devez déterminer la structure JSON des données retournées. Pour les entrées et les sorties de votre script, envisagez de faire `interface` un pour correspondre aux structures JSON nécessaires. Cela donne au script plus de sécurité de type. Vous pouvez voir un exemple de cela dans [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitations avec les appels externes de Office Scripts

* Il n’y a aucun moyen de se connecter ou d’utiliser le type de flux d’authentification OAuth2. Toutes les clés et informations d’identification doivent être codées en dur (ou lues à partir d’une autre source).
* Il n’existe pas d’infrastructure pour stocker les informations d’identification et les clés de l’API. Cela devra être géré par l’utilisateur.
* Documentez les cookies et `localStorage` les objets ne sont pas pris en `sessionStorage` charge. 
* Les appels externes peuvent entraîner l’exposition de données sensibles à des paramètres indésirables ou l’entrée de données externes dans des cahiers de travail internes. Votre administrateur peut établir une protection pare-feu contre de tels appels. Assurez-vous de vérifier auprès des politiques locales avant de vous fier aux appels externes.
* Assurez-vous de vérifier la quantité de débit de données avant de prendre une dépendance. Par exemple, tirer vers le bas l’ensemble de données externes peut ne pas être la meilleure option et au lieu de pagination devrait être utilisé pour obtenir des données en morceaux.

## <a name="retrieve-information-with-fetch"></a>Récupérer des informations avec `fetch`

[L’API fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) récupère des informations auprès de services externes. Il s’agit `async` d’une API, vous devez donc ajuster `main` la signature de votre script. Faire la `main` fonction et lui faire retourner un `async` `Promise<void>` . Vous devez également être sûr de `await` `fetch` `json` l’appel et de la récupération. Cela garantit que ces opérations sont terminées avant la fin du script.

Toutes les données JSON récupérées par `fetch` doit correspondre à une interface définie dans le script. La valeur retournée doit être attribuée à un type spécifique car [Office scripts ne supportent pas le `any` type](typescript-restrictions.md#no-any-type-in-office-scripts). Vous devez vous référer à la documentation de votre service pour voir quels sont les noms et les types de propriétés retournées. Ensuite, ajoutez l’interface ou les interfaces correspondantes à votre script.

Le script suivant utilise `fetch` pour récupérer les données JSON du serveur de test dans l’URL donnée. Notez `JSONData` l’interface pour stocker les données comme un type correspondant.

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### <a name="other-fetch-samples"></a>Autres `fetch` échantillons

* [L’exemple Utiliser des appels externes d’extraction Office scripts](../resources/samples/external-fetch-calls.md) montre comment obtenir des informations de base sur les référentiels de GitHub utilisateur.
* Le [scénario de l’exemple Office Scripts : Graph données au niveau de l’eau de la NOAA démontrent](../resources/scenarios/noaa-data-fetch.md) que la commande d’extraction est utilisée pour récupérer les enregistrements de la base de données Tides and Currents de la National Oceanic and Atmospheric Administration.

## <a name="external-calls-from-power-automate"></a>Appels externes de Power Automate

Tout appel API externe échoue lorsqu’un script est exécuté avec Power Automate. Il s’agit d’une différence comportementale entre l’exécution d’un script à travers Excel’application et à travers Power Automate. Assurez-vous de vérifier vos scripts pour de telles références avant de les construire dans un flux.

Vous devrez utiliser HTTP avec [Azure AD ou d’autres](/connectors/webcontents/) actions équivalentes pour extraire des données ou les pousser vers un service externe.

> [!WARNING]
> Les appels externes effectués par l’intermédiaire Power Automate [Excel connecteur en ligne](/connectors/excelonlinebusiness) échouent afin d’aider à maintenir les politiques existantes de prévention des pertes de données. Toutefois, les scripts qui sont exécutés à travers Power Automate sont effectués en dehors de votre organisation, et en dehors des pare-feu de votre organisation. Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation Office scripts. Votre administrateur peut désactiver le connecteur Excel En ligne en Power Automate ou désactiver les scripts Office pour Excel sur le Web à travers [les contrôles d’administrateur Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Voir aussi

* [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
* [Utiliser les appels externes de récupération à Office Scripts](../resources/samples/external-fetch-calls.md)
* [Office Scénario de l’échantillon scripts : Graph données sur le niveau de l’eau de la NOAA](../resources/scenarios/noaa-data-fetch.md)
