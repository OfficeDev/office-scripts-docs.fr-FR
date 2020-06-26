---
title: Différences entre les scripts Office et les compléments Office
description: Les différences de comportement et d’API entre les scripts Office et les compléments Office.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: fc2029780190672c633e00e26f44273e4311c754
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878660"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Différences entre les scripts Office et les compléments Office

Les compléments Office et les scripts Office ont beaucoup de choses en commun. Elles offrent toutes deux le contrôle automatique d’un classeur Excel d’une API JavaScript. Toutefois, les API de scripts Office sont une version spécifique et synchrone de l’API JavaScript pour Office.

![Diagramme à quatre quadrants montrant les zones ciblées pour différentes solutions d’extensibilité Office. Les scripts Office et les compléments Office Web sont centrés sur le Web et la collaboration, mais les scripts Office répondent aux utilisateurs finaux (tandis que les compléments Web Office ciblent les développeurs professionnels).)](../images/office-programmability-diagram.png)

Les scripts Office sont exécutés jusqu’à la fin avec une pression manuelle ou une étape de l' [automate d’alimentation](https://flow.microsoft.com/), tandis que les compléments Office persistent lorsque leurs volets Office sont ouverts. Cela signifie que les compléments peuvent conserver l’État pendant une session, tandis que les scripts Office ne gèrent pas un état interne entre les exécutions. Si vous constatez que votre extension Excel doit dépasser les fonctionnalités de la plateforme de script, consultez la [documentation relative aux compléments Office](/office/dev/add-ins) pour en savoir plus sur les compléments Office.

Le reste de cet article décrit les principales différences entre les compléments Office et les scripts Office.

## <a name="platform-support"></a>Prise en charge de la plateforme

Les compléments Office sont multiplateformes. Elles fonctionnent sur des plateformes de bureau Windows, Mac, iOS et Web et fournissent la même expérience sur chacun d’eux. Toutes les exceptions sont indiquées dans la documentation de l’API individuelle.

Les scripts Office sont actuellement uniquement pris en charge par Excel sur le Web. Toutes les opérations d’enregistrement, de modification et d’exécution sont réalisées sur la plateforme Web.

## <a name="apis"></a>API

Il n’existe pas de version synchrone des API JavaScript pour Office pour les compléments Office. Les API de scripts Office standard sont propres à la plateforme et présentent de nombreuses optimisations et altérations pour éviter l’utilisation du `load` / `sync` paradigme.

Certaines [API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview) sont compatibles avec les [API Async de scripts Office](../develop/excel-async-model.md). Certains exemples et blocs de code de complément peuvent être transférés vers `Excel.run` des blocs avec une traduction minimale. Bien que les deux plateformes partagent la fonctionnalité, il existe des lacunes. Les deux principaux ensembles d’API que les compléments Office ont, mais les scripts Office ne sont pas des événements et les API communes.

### <a name="events"></a>Événements

Les scripts Office ne prennent pas en charge les [événements](/office/dev/add-ins/excel/excel-add-ins-events). Chaque script exécute le code dans une `main` méthode unique, puis se termine. Il ne réactive pas lorsque des événements sont déclenchés et, par conséquent, ne peut pas enregistrer d’événements.

### <a name="common-apis"></a>API courantes

Les scripts Office ne peuvent pas utiliser des [API communes](/javascript/api/office). Si vous avez besoin d’une authentification, de fenêtres de boîtes de dialogue ou d’autres fonctionnalités qui sont uniquement prises en charge par des API communes, vous aurez probablement besoin de créer un complément Office au lieu d’un script Office.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Différences entre les scripts Office et les macros VBA](vba-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Créer un complément de volet de tâches Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
