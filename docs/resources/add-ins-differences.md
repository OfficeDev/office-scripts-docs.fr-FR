---
title: Différences entre les scripts Office et les compléments Office
description: Les différences de comportement et d’API entre Office scripts et Office des modules.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 45993d08d85cfceb299216dddbe2e7da9fd2e404
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232633"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Différences entre les scripts Office et les compléments Office

Office Les add-ins et Office scripts ont beaucoup en commun. Ils offrent tous deux un contrôle automatisé d’un Excel de travail une API JavaScript. Toutefois, les API Office Scripts sont une version spécialisée et synchrone de l Office API JavaScript.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les zones de mise au point pour Office solutions d’extensibilité. Les scripts Office et les applications web Office sont axés sur le web et la collaboration, mais les scripts Office sont pris en compte par les utilisateurs finaux (tandis que les Office web s’adressent aux développeurs professionnels)":::

Office Les scripts s’exécutent jusqu’à la fin avec un bouton manuel ou une étape dans [Power Automate](https://flow.microsoft.com/), tandis que les Office sont persistants pendant que leurs volets Des tâches sont ouverts. Cela signifie que les add-ins peuvent maintenir l’état pendant une session, tandis que Office Scripts ne conservent pas un état interne entre les séquences. Si vous constatez que votre extension Excel doit dépasser les [fonctionnalités](/office/dev/add-ins) de la plateforme de script, consultez la documentation des Office pour en savoir plus sur les Office de script.

Le reste de cet article décrit les principales différences entre les Office et Office scripts.

## <a name="platform-support"></a>Prise en charge de la plateforme

Office Les add-ins sont sur plusieurs plateformes. Ils fonctionnent sur Windows de bureau, Mac, iOS et les plateformes web et offrent la même expérience sur chacune d’elles. Toute exception à cette règle est notée dans la documentation de l’API individuelle.

Office Les scripts sont actuellement uniquement pris en charge par les Excel sur le Web. L’enregistrement, la modification et l’exécution s’exécutent sur la plateforme web.

## <a name="apis"></a>API

Il n’existe aucune version synchrone des API JavaScript Office pour les Office de recherche. Les API Office scripts standard sont propres à la plateforme et ont de nombreuses optimisations et modifications pour éviter l’utilisation du `load` / `sync` paradigme.

Certaines des API [JavaScript Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true) sont compatibles avec les API [async Office scripts.](../develop/excel-async-model.md) Certains exemples et blocs de code de add-in peuvent être portés vers des `Excel.run` blocs avec une traduction minimale. Bien que les deux plateformes partagent des fonctionnalités, il existe des lacunes. Les deux principaux ensembles d’API Office les Office les scripts ne sont pas des événements et les API communes.

### <a name="events"></a>Événements

Office Les scripts ne sont pas en charge les [événements](/office/dev/add-ins/excel/excel-add-ins-events). Chaque script exécute le code dans une seule `main` méthode, puis se termine. Elle ne se réactive pas lorsque des événements sont déclenchés et, par conséquent, ne peut pas enregistrer d’événements.

### <a name="common-apis"></a>API courantes

Office Les scripts ne peuvent pas utiliser [les API communes.](/javascript/api/office) Si vous avez besoin d’une authentification, de fenêtres de boîte de dialogue ou d’autres fonctionnalités uniquement pris en charge par les API communes, vous devrez probablement créer un add-in Office au lieu d’un script Office.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Différences entre Office scripts et les macros VBA](vba-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Créer un complément de volet de tâches Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
