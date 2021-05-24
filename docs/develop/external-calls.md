---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Prise en charge et conseils pour effectuer des appels d’API externes dans Office Script.
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

Les auteurs de scripts ne doivent pas s’attendre à un comportement cohérent lors de l’utilisation [d’API](https://developer.mozilla.org/docs/Web/API) externes lors de la phase de prévisualisation de la plateforme. En tant que tel, ne comptez pas sur les API externes pour les scénarios de script critiques.

Les appels aux API externes peuvent uniquement être effectués via l’application Excel, et non Power Automate [dans des circonstances normales.](#external-calls-from-power-automate)

> [!CAUTION]
> Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables. Votre administrateur peut établir une protection pare-feu contre ces appels.

## <a name="configure-your-script-for-external-calls"></a>Configurer votre script pour les appels externes

Les appels externes [sont asynchrones](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) et nécessitent que votre script soit marqué comme `async` . Ajoutez `async` le préfixe à votre fonction et lui `main` renvoyez un , comme illustré ici `Promise` :

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Les scripts qui retournent d’autres informations peuvent `Promise` renvoyer un type de ce type. Par exemple, si votre script doit renvoyer un `Employee` objet, la signature de retour sera `: Promise <Employee>`

Vous devez découvrir les interfaces du service externe pour appeler ce service. Si vous utilisez ou des API REST, vous devez déterminer la `fetch` structure JSON des données renvoyées. [](https://wikipedia.org/wiki/Representational_state_transfer) Pour l’entrée et la sortie de votre script, envisagez d’effectuer une correspondance avec `interface` les structures JSON nécessaires. Cela permet au script d’améliorer la sécurité des types. Vous pouvez en voir un exemple dans [l’utilisation de la récupération à partir Office scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitations avec les appels externes de Office Scripts

* Il n’existe aucun moyen de se connecter ou d’utiliser le type de flux d’authentification OAuth2. Toutes les clés et informations d’identification doivent être codées en dur (ou lues à partir d’une autre source).
* Il n’existe aucune infrastructure pour stocker les informations d’identification et les clés d’API. Il devra être géré par l’utilisateur.
* Les cookies de document `localStorage` et les objets ne sont pas pris en `sessionStorage` charge. 
* Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables ou des données externes à mettre dans des workbooks internes. Votre administrateur peut établir une protection pare-feu contre ces appels. Veillez à vérifier les stratégies locales avant de vous appuyer sur des appels externes.
* Veillez à vérifier la quantité de débit de données avant de prendre une dépendance. Par exemple, il est possible que le fait d’retirer l’intégralité du jeu de données externe ne soit pas la meilleure option et que la pagination soit utilisée pour obtenir des données par blocs.

## <a name="retrieve-information-with-fetch"></a>Récupérer des informations avec `fetch`

[L’API de](https://developer.mozilla.org/docs/Web/API/Fetch_API) récupération récupère des informations à partir de services externes. Il s’agit `async` d’une API, vous devez donc ajuster la `main` signature de votre script. Make the `main` function and have it return a `async` `Promise<void>` . Vous devez également être sûr de `await` `fetch` l’appel et de la `json` récupération. Cela garantit que ces opérations sont terminées avant la fin du script.

Toutes les données JSON récupérées par `fetch` doivent correspondre à une interface définie dans le script. La valeur renvoyée doit être affectée à un type spécifique, [car les scripts Office ne le supportent `any` pas.](typescript-restrictions.md#no-any-type-in-office-scripts) Vous devez consulter la documentation de votre service pour voir les noms et les types des propriétés renvoyées. Ensuite, ajoutez l’interface ou les interfaces correspondantes à votre script.

Le script suivant utilise `fetch` pour récupérer les données JSON du serveur de test dans l’URL donnée. Notez `JSONData` l’interface pour stocker les données en tant que type correspondant.

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

### <a name="other-fetch-samples"></a>Autres `fetch` exemples

* L’exemple Utiliser des appels de récupération externe [dans Office Scripts](../resources/samples/external-fetch-calls.md) montre comment obtenir des informations de base sur les référentiels GitHub d’un utilisateur.
* L’exemple de scénario [Office Scripts](../resources/scenarios/noaa-data-fetch.md) : Graph données au niveau de l’eau de la NOAA illustre la commande de récupération utilisée pour extraire des enregistrements de la base de données Archives et courants de l’administration nationale.

## <a name="external-calls-from-power-automate"></a>Appels externes de Power Automate

Tout appel d’API externe échoue lorsqu’un script est exécuté avec Power Automate. Il s’agit d’une différence de comportement entre l’exécution d’un script via l Excel’application Power Automate. Veillez à vérifier si vos scripts sont de telles références avant de les créer dans un flux.

Vous devez utiliser HTTP avec [Azure AD](/connectors/webcontents/) ou d’autres actions équivalentes pour tirer des données ou les pousser vers un service externe.

> [!WARNING]
> Les appels externes effectués via le connecteur Power Automate [Excel Online](/connectors/excelonlinebusiness) échouent pour aider à respecter les stratégies de protection contre la perte de données existantes. Toutefois, les scripts exécutés par Power Automate sont effectués en dehors de votre organisation et en dehors des pare-feu de votre organisation. Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation Office scripts. Votre administrateur peut désactiver le connecteur Excel Online dans Power Automate ou désactiver les scripts Office pour Excel sur le Web via les contrôles d’administrateur [Office Scripts.](/microsoft-365/admin/manage/manage-office-scripts-settings)

## <a name="see-also"></a>Voir aussi

* [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
* [Utiliser les appels externes de récupération à Office Scripts](../resources/samples/external-fetch-calls.md)
* [Office Exemple de scénario de scripts : Graph données de niveau d’eau à partir de la NOAA](../resources/scenarios/noaa-data-fetch.md)
