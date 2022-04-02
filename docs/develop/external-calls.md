---
title: Prise en charge des appels d’API externes dans Scripts Office
description: Prise en charge et conseils pour effectuer des appels d’API externes dans Office script.
ms.date: 05/21/2021
ms.localizationpriority: medium
ms.openlocfilehash: abcd548c9b62ce9bd5c40866915ae50a6d1cc5be
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585736"
---
# <a name="external-api-call-support-in-office-scripts"></a>Prise en charge des appels d’API externes dans Scripts Office

Les scripts prendre en charge les appels à des services externes. Utilisez ces services pour fournir des données et d’autres informations à votre workbook.

> [!CAUTION]
> Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables. Votre administrateur peut établir une protection pare-feu contre ces appels.

> [!IMPORTANT]
> Les appels aux API externes peuvent uniquement être effectués via l’application Excel, et non par Power Automate [dans des circonstances normales](#external-calls-from-power-automate).

## <a name="configure-your-script-for-external-calls"></a>Configurer votre script pour les appels externes

Les appels externes [sont asynchrones](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) et nécessitent que votre script soit marqué comme `async`. Ajoutez le `async` préfixe à votre `main` fonction et lui renvoyez un `Promise`, comme illustré ici :

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Les scripts qui retournent d’autres informations peuvent renvoyer un `Promise` type de ce type. Par exemple, si votre script doit renvoyer un `Employee` objet, la signature de retour sera `: Promise <Employee>`

Vous devez découvrir les interfaces du service externe pour effectuer des appels à ce service. Si vous utilisez ou `fetch` [des API REST](https://wikipedia.org/wiki/Representational_state_transfer), vous devez déterminer la structure JSON des données renvoyées. Pour l’entrée et la sortie de votre script, envisagez d’effectuer une `interface` correspondance avec les structures JSON nécessaires. Cela permet au script d’améliorer la sécurité des types. Vous pouvez en voir un exemple dans [l’utilisation de la Office scripts](../resources/samples/external-fetch-calls.md).

### <a name="limitations-with-external-calls-from-office-scripts"></a>Limitations avec les appels externes de Office Scripts

* Il n’existe aucun moyen de se connecter ou d’utiliser le type de flux d’authentification OAuth2. Toutes les clés et informations d’identification doivent être codées en dur (ou lues à partir d’une autre source).
* Il n’existe aucune infrastructure pour stocker les informations d’identification et les clés d’API. Il devra être géré par l’utilisateur.
* Les cookies de document et `localStorage`les objets `sessionStorage` ne sont pas pris en charge.
* Les appels externes peuvent entraîner l’exposition de données sensibles à des points de terminaison indésirables ou des données externes à entrer dans des workbooks internes. Votre administrateur peut établir une protection pare-feu contre ces appels. Veillez à vérifier les stratégies locales avant de vous appuyer sur des appels externes.
* Veillez à vérifier la quantité de débit de données avant de prendre une dépendance. Par exemple, il est possible que le fait d’retirer l’intégralité du jeu de données externe ne soit pas la meilleure option et que la pagination soit utilisée pour obtenir des données en blocs.

## <a name="retrieve-information-with-fetch"></a>Récupérer des informations avec `fetch`

[L’API de](https://developer.mozilla.org/docs/Web/API/Fetch_API) récupération récupère des informations à partir de services externes. Il s’agit d’une `async` API, vous devez donc ajuster la `main` signature de votre script. Rendre la `main` fonction `async`. Vous devez également être sûr de l’appel `await` `fetch` et `json` de la récupération. Cela garantit que ces opérations sont terminées avant la fin du script.

Toutes les données JSON récupérées par `fetch` doivent correspondre à une interface définie dans le script. La valeur renvoyée doit être affectée à un type spécifique, car Office [scripts ne le prisent pas en `any` charge](typescript-restrictions.md#no-any-type-in-office-scripts). Vous devez consulter la documentation de votre service pour voir les noms et les types des propriétés renvoyées. Ensuite, ajoutez l’interface ou les interfaces correspondantes à votre script.

Le script suivant utilise pour `fetch` récupérer les données JSON du serveur de test dans l’URL donnée. Notez l’interface `JSONData` pour stocker les données en tant que type correspondant.

```TypeScript
async function main(workbook: ExcelScript.Workbook){
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
* L’exemple de scénario [Office Scripts : Graph](../resources/scenarios/noaa-data-fetch.md) données au niveau de l’eau de la NOAA illustre la commande de récupération utilisée pour extraire des enregistrements de la base de données Archives et courants de l’administration nationale.

## <a name="external-calls-from-power-automate"></a>Appels externes de Power Automate

Tout appel d’API externe échoue lorsqu’un script est exécuté avec Power Automate. Il s’agit d’une différence comportementale entre l’exécution d’un script via Excel application et par le biais Power Automate. Veillez à vérifier si vos scripts sont de telles références avant de les créer dans un flux.

Vous devez utiliser [HTTP](/connectors/webcontents/) avec des Azure AD ou d’autres actions équivalentes pour tirer des données ou les pousser vers un service externe.

> [!WARNING]
> Les appels externes effectués via le connecteur Power Automate [Excel Online](/connectors/excelonlinebusiness) échouent afin de contribuer à la protection contre la perte de données existantes. Toutefois, les scripts exécutés par Power Automate sont effectués en dehors de votre organisation et en dehors des pare-feu de votre organisation. Pour une protection supplémentaire contre les utilisateurs malveillants dans cet environnement externe, votre administrateur peut contrôler l’utilisation Office scripts. Votre administrateur peut désactiver le connecteur Excel Online dans Power Automate ou désactiver Office Scripts pour Excel sur le Web via les contrôles d’administrateur [Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="see-also"></a>Voir aussi

* [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
* [Utiliser les appels externes de récupération dans les scripts Office](../resources/samples/external-fetch-calls.md)
* [Office exemple de scripts : Graph données de niveau d’eau à partir de la NOAA](../resources/scenarios/noaa-data-fetch.md)
