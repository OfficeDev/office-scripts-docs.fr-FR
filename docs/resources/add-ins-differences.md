---
title: Différences entre les scripts Office et les compléments Office
description: Les différences de comportement et d’API entre les scripts Office et les compléments Office.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978714"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Différences entre les scripts Office et les compléments Office

Les compléments Office et les scripts Office ont beaucoup de choses en commun. Ils proposent tous les deux un contrôle automatique d’un classeur Excel via l' `Excel` espace de noms de l’API JavaScript pour Office. Toutefois, les scripts Office sont plus limités dans leur étendue.

![Diagramme à quatre quadrants montrant les zones ciblées pour différentes solutions d’extensibilité Office. Les scripts Office et les compléments Office Web sont centrés sur le Web et la collaboration, mais les scripts Office répondent aux utilisateurs finaux (tandis que les compléments Web Office ciblent les développeurs professionnels).)](../images/office-programmability-diagram.png)

Les scripts Office sont exécutés jusqu’à la fin avec une pression manuelle ou une étape de l' [automate d’alimentation](https://flow.microsoft.com/), tandis que les compléments Office persistent lorsque leurs volets Office sont ouverts. Cela signifie que les compléments peuvent conserver l’État pendant une session, tandis que les scripts Office ne gèrent pas un état interne entre les exécutions. Si vous constatez que votre extension Excel doit dépasser les fonctionnalités de la plateforme de script, consultez la [documentation relative aux compléments Office](/office/dev/add-ins) pour en savoir plus sur les compléments Office.

Le reste de cet article décrit les principales différences entre les compléments Office et les scripts Office.

## <a name="platform-support"></a>Prise en charge de la plateforme

Les compléments Office sont multiplateformes. Elles fonctionnent sur des plateformes de bureau Windows, Mac, iOS et Web et fournissent la même expérience sur chacun d’eux. Toutes les exceptions sont indiquées dans la documentation de l’API individuelle.

Les scripts Office sont actuellement uniquement pris en charge par Excel sur le Web. Toutes les opérations d’enregistrement, de modification et d’exécution sont réalisées sur la plateforme Web.

## <a name="apis"></a>API

Les scripts Office prennent en charge la plupart des API JavaScript pour Excel, ce qui signifie qu’il existe un grand nombre de fonctionnalités qui se chevauchent entre les deux plateformes. Il existe deux exceptions : les événements et les API communes.

### <a name="events"></a>Événements

Les scripts Office ne prennent pas en charge les [événements](/office/dev/add-ins/excel/excel-add-ins-events). Chaque script exécute le code dans une méthode `main` unique, puis se termine. Il ne réactive pas lorsque des événements sont déclenchés et, par conséquent, ne peut pas enregistrer d’événements.

### <a name="common-apis"></a>API communes

Les scripts Office ne peuvent pas utiliser des [API communes](/javascript/api/office). Si vous avez besoin d’une authentification, de fenêtres de boîtes de dialogue ou d’autres fonctionnalités qui sont uniquement prises en charge par des API communes, vous aurez probablement besoin de créer un complément Office au lieu d’un script Office.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Différences entre les scripts Office et les macros VBA](vba-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Créer un complément de volet de tâches Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
