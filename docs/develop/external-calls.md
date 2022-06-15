---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Prise en charge et conseils pour effectuer des appels d’API externes dans un script Office.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: b847400893184533c250ab99b640563ff0cbdb3e
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088042"
---
# <a name="external-api-call-support-in-office-scripts"></a>Prise en charge des appels d’API externes dans Scripts Office

Les scripts prennent en charge les appels aux services externes. Utilisez ces services pour fournir des données et d’autres informations à votre classeur.

> [!CAUTION]
> Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables. Votre administrateur peut établir une protection pare-feu contre ces appels.

> [!IMPORTANT]
> Les appels aux API externes ne peuvent être effectués que par le biais de l’application Excel, et non par Power Automate [dans des circonstances normales](#external-calls-from-power-automate).

## <a name="configure-your-script-for-external-calls"></a>Configurer votre script pour les appels externes

Les appels externes sont [asynchrones](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) et nécessitent que votre script soit marqué comme `async`. Ajoutez le `async` préfixe à votre `main` fonction et faites-la retourner, `Promise`comme illustré ici :

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Les scripts qui retournent d’autres informations peuvent retourner un `Promise` de ce type. Par exemple, si votre script doit retourner un `Employee` objet, la signature de retour est `: Promise <Employee>`

Vous devez apprendre les interfaces du service externe pour effectuer des appels à ce service. Si vous utilisez ou si vous utilisez `fetch` [des API REST](https://wikipedia.org/wiki/Representational_state_transfer), vous devez déterminer la structure JSON des données retournées. Pour l’entrée et la sortie à partir de votre script, envisagez de faire correspondre `interface` les structures JSON nécessaires. Cela permet au script d’être plus sûr du type. Vous pouvez en voir un exemple dans [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitations avec les appels externes à partir de scripts Office

* Il n’existe aucun moyen de se connecter ou d’utiliser le type de flux d’authentification OAuth2. Toutes les clés et informations d’identification doivent être codées en dur (ou lues à partir d’une autre source).
* Il n’existe aucune infrastructure pour stocker les informations d’identification et les clés de l’API. Cela doit être géré par l’utilisateur.
* Les cookies de document `localStorage`et `sessionStorage` les objets ne sont pas pris en charge.
* Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables ou l’insertion de données externes dans des classeurs internes. Votre administrateur peut établir une protection pare-feu contre ces appels. Veillez à vérifier auprès des stratégies locales avant de vous appuyer sur des appels externes.
* Veillez à vérifier la quantité de débit des données avant de prendre une dépendance. Par exemple, l’extraction de l’intégralité du jeu de données externe peut ne pas être la meilleure option. À la place, la pagination doit être utilisée pour obtenir des données en blocs.

## <a name="retrieve-information-with-fetch"></a>Récupérer des informations avec `fetch`

[L’API fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) récupère des informations à partir de services externes. Il s’agit d’une `async` API. Vous devez donc ajuster la `main` signature de votre script. Créer la `main` fonction `async`. Vous devez également être sûr de `await` l’appel et `json` de la `fetch` récupération. Cela garantit que ces opérations se terminent avant la fin du script.

Toutes les données JSON récupérées `fetch` doivent correspondre à une interface définie dans le script. La valeur retournée doit être affectée à un type spécifique, car [Office scripts ne prennent pas en charge le `any` type](typescript-restrictions.md#no-any-type-in-office-scripts). Vous devez consulter la documentation de votre service pour voir quels sont les noms et les types des propriétés retournées. Ajoutez ensuite l’interface ou les interfaces correspondantes à votre script.

Le script suivant permet de récupérer des `fetch` données JSON à partir du serveur de test dans l’URL donnée. Notez l’interface `JSONData` pour stocker les données en tant que type correspondant.

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
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

* L’exemple [Utiliser des appels de récupération externe dans Office Scripts](../resources/samples/external-fetch-calls.md) montre comment obtenir des informations de base sur les référentiels GitHub d’un utilisateur.
* [L’exemple de scénario Office Scripts : Graph données au niveau de l’eau de la NOAA](../resources/scenarios/noaa-data-fetch.md) illustrent la commande fetch utilisée pour récupérer les enregistrements de la base de données Tides and Currents de l’Administration nationale océanique et atmosphérique.

## <a name="external-calls-from-power-automate"></a>Appels externes à partir de Power Automate

Tout appel d’API externe échoue lorsqu’un script est exécuté avec Power Automate. Il s’agit d’une différence comportementale entre l’exécution d’un script via l’application Excel et Power Automate. Veillez à rechercher ces références dans vos scripts avant de les intégrer dans un flux.

Vous devez utiliser [HTTP avec Azure AD](/connectors/webcontents/) ou d’autres actions équivalentes pour extraire ou envoyer (push) des données à un service externe.

> [!WARNING]
> Les appels externes effectués via le [connecteur Power Automate Excel Online](/connectors/excelonlinebusiness) échouent pour aider à respecter les stratégies de protection contre la perte de données existantes. Toutefois, les scripts exécutés via Power Automate le sont en dehors de votre organisation et en dehors des pare-feu de votre organisation. Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation de scripts Office. Votre administrateur peut désactiver le connecteur Excel Online dans Power Automate ou désactiver Office Scripts pour Excel sur le Web via les [contrôles administrateur Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Voir aussi

* [Utiliser JSON pour transmettre des données vers et depuis Office Scripts](use-json.md)
* [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
* [Utiliser les appels externes de récupération dans les scripts Office](../resources/samples/external-fetch-calls.md)
* [exemple de scénario Office Scripts : Graph données de niveau de l’eau à partir de la NOAA](../resources/scenarios/noaa-data-fetch.md)
