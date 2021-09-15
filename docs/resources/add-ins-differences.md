---
title: Différences entre les scripts Office et les compléments Office
description: Les différences de comportement et d’API entre Office scripts et Office des modules.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7b199e8f3acdbe753fcaa2d1f4b6b5f11998b52b
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59328099"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Différences entre les scripts Office et les compléments Office

Comprendre les différences entre Office scripts et les Office pour savoir quand les utiliser. Office Les scripts sont conçus pour être créés rapidement par toute personne cherchant à améliorer son flux de travail. Office Les add-ins s’intègrent à Office’interface utilisateur pour une expérience plus interactive via les boutons du ruban et les volets Des tâches. Office Les add-ins peuvent également développer des fonctions intégrées Excel en fournissant des fonctions personnalisées.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les zones de mise au point pour Office solutions d’extensibilité. Les scripts Office et les applications web Office sont axés sur le web et la collaboration, mais les scripts Office sont pris en compte par les utilisateurs finaux (tandis que les Office web add-ins ciblent les développeurs professionnels).":::

Office Les scripts s’exécutent jusqu’à la fin avec un bouton manuel ou à l’étape de [Power Automate](https://flow.microsoft.com/), tandis que les Office se poursuivent en fonction de la façon dont ils sont configurés. Par exemple, vous pouvez configurer un Office pour qu’il continue à s’exécute même lorsque son volet Des tâches est fermé. Cela signifie que les Office de gestion conservent l’état au cours d’une session, tandis que Office Scripts ne conservent pas d’état interne entre les séquences. Si la solution que vous construisez nécessite un état de mise à jour, consultez la [documentation](/office/dev/add-ins) des Office pour en savoir plus sur les Office de développement.

Le reste de cet article décrit les principales différences entre les Office et Office scripts.

## <a name="platform-support"></a>Prise en charge de la plateforme

Office Les add-ins sont sur plusieurs plateformes. Ils fonctionnent sur Windows de bureau, Mac, iOS et les plateformes web et offrent la même expérience sur chacune d’elles. Toute exception à cette règle est notée dans la documentation de l’API individuelle.

Office Les scripts sont actuellement uniquement pris en charge par les Excel sur le Web. L’enregistrement, la modification et l’exécution s’exécutent sur la plateforme web.

## <a name="apis"></a>API

Bien que les OFFICE JavaScript pour les Office et les API Office Scripts partagent certaines fonctionnalités, ce sont des plateformes différentes. Les API Office scripts sont un sous-ensemble optimisé et synchrone du modèle d Excel API JavaScript. La principale différence est l’utilisation du `load` / `sync` paradigme avec les applications. En outre, les compléments offrent des API pour les événements et un ensemble plus large de fonctionnalités en dehors des Excel, appelés API communes.

### <a name="events"></a>Events

Office Les scripts ne supportent pas les événements au niveau du [workbook.](/office/dev/add-ins/excel/excel-add-ins-events) Les scripts sont déclenchés par  les utilisateurs qui sélectionnent le bouton Exécuter pour un script ou par Power Automate. Chaque script exécute le code dans une seule `main` méthode, puis se termine.

### <a name="common-apis"></a>API courantes

Office Les scripts ne peuvent pas utiliser [les API communes.](/javascript/api/office) Si vous avez besoin d’une authentification, de fenêtres de boîte de dialogue ou d’autres fonctionnalités uniquement pris en charge par les API communes, vous devrez probablement créer un add-in Office au lieu d’un script Office.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Différences entre Office scripts et les macros VBA](vba-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Créer un complément de volet de tâches Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
