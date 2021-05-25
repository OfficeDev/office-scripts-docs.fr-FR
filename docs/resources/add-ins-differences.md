---
title: Différences entre les scripts Office et les compléments Office
description: Les différences de comportement et d’API entre Office scripts et Office des modules.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 5c30406867da05952dedda684f765df5e7a7e53f
ms.sourcegitcommit: 09d8859d5269ada8f1d0e141f6b5a4f96d95a739
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2021
ms.locfileid: "52631677"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Différences entre les scripts Office et les compléments Office

Office Les scripts de Office et de Office ont beaucoup de points communs. Ils offrent tous deux un contrôle automatisé d’un Excel de travail une API JavaScript. Toutefois, les API Office Scripts sont une version spécialisée et synchrone de l Office API JavaScript.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les zones de mise au point pour Office solutions d’extensibilité. Les scripts Office et les applications web Office sont axés sur le web et la collaboration, mais les scripts Office sont pris en compte par les utilisateurs finaux (tandis que les Office web s’adressent aux développeurs professionnels)":::

Office Les scripts s’exécutent jusqu’à la fin avec un bouton manuel ou une étape dans [Power Automate](https://flow.microsoft.com/), tandis que les Office sont persistants pendant que leurs volets Des tâches sont ouverts. Cela signifie que les add-ins peuvent maintenir l’état pendant une session, tandis que Office Scripts ne conservent pas un état interne entre les séquences. Si vous constatez que votre extension Excel doit dépasser les [fonctionnalités](/office/dev/add-ins) de la plateforme de script, consultez la documentation des Office pour en savoir plus sur les Office de script.

Le reste de cet article décrit les principales différences entre les Office et Office scripts.

## <a name="platform-support"></a>Prise en charge de la plateforme

Office Les add-ins sont sur plusieurs plateformes. Ils fonctionnent sur Windows de bureau, Mac, iOS et les plateformes web et offrent la même expérience sur chacune d’elles. Toute exception à cette règle est notée dans la documentation de l’API individuelle.

Office Les scripts sont actuellement uniquement pris en charge par les Excel sur le Web. L’enregistrement, la modification et l’exécution s’exécutent sur la plateforme web.

## <a name="apis"></a>API

Bien que les OFFICE JavaScript pour les Office et les API Office Scripts partagent certaines fonctionnalités, ce sont des plateformes différentes. Les API Office Scripts sont une version optimisée et synchrone du modèle d’API JavaScript Excel. La principale différence est l’utilisation du `load` / `sync` paradigme avec les applications. En outre, les compléments offrent des API pour les événements et un ensemble plus large de fonctionnalités en dehors des Excel, appelés API communes.

### <a name="events"></a>Événements

Office Les scripts ne sont pas en charge les [événements](/office/dev/add-ins/excel/excel-add-ins-events). Chaque script exécute le code dans une seule `main` méthode, puis se termine. Elle ne se réactive pas lorsque des événements sont déclenchés et, par conséquent, ne peut pas enregistrer d’événements.

### <a name="common-apis"></a>API courantes

Office Les scripts ne peuvent pas utiliser [les API communes.](/javascript/api/office) Si vous avez besoin d’une authentification, de fenêtres de boîte de dialogue ou d’autres fonctionnalités uniquement pris en charge par les API communes, vous devrez probablement créer un add-in Office au lieu d’un script Office.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Différences entre Office scripts et les macros VBA](vba-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Créer un complément de volet de tâches Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
