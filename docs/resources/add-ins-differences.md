---
title: Différences entre les scripts Office et les compléments Office
description: Comportement et différences d’API entre les scripts Office et les compléments Office.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: a3df4daf04f963598d2cb31f82dd2c1c9923fdc8
ms.sourcegitcommit: 33fe0f6807daefb16b148fd73c863de101f47cea
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/08/2022
ms.locfileid: "67281909"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Différences entre les scripts Office et les compléments Office

Découvrez les différences entre les scripts Office et les compléments Office pour savoir quand utiliser chacun d’eux. Les scripts Office sont conçus pour être rapidement créés par toute personne cherchant à améliorer son flux de travail. Les compléments Office s’intègrent à l’interface utilisateur Office pour une expérience plus interactive via les boutons du ruban et les volets Office. Les compléments Office peuvent également développer des fonctions Excel intégrées en fournissant des fonctions personnalisées.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les zones de focus pour différentes solutions d’extensibilité Office. Les scripts Office et les compléments Web Office sont axés sur le web et la collaboration, mais les scripts Office s’adressent aux utilisateurs finaux (tandis que les compléments Web Office ciblent les développeurs professionnels).":::

Les scripts Office s’exécutent jusqu’à la fin avec un appui manuel sur un bouton ou en tant qu’étape dans [Power Automate](https://flow.microsoft.com/), tandis que les compléments Office continuent à s’exécuter en fonction de la façon dont ils sont configurés. Par exemple, vous pouvez configurer un complément Office pour continuer à s’exécuter même lorsque son volet Office est fermé. Cela signifie que les compléments Office conservent l’état pendant une session, tandis que les scripts Office ne conservent pas d’état interne entre les exécutions. Si la solution que vous créez nécessite un état maintenu, vous devez consulter la [documentation des compléments Office](/office/dev/add-ins) pour en savoir plus sur les compléments Office.

Le reste de cet article décrit les principales différences entre les compléments Office et les scripts Office.

## <a name="platform-support"></a>Prise en charge de la plateforme

Les compléments Office sont multiplateformes. Ils fonctionnent sur les plateformes de bureau, Mac, iOS et web Windows et offrent la même expérience sur chacune d’elles. Toute exception à cette règle est indiquée dans la documentation de l’API individuelle.

Les scripts Office ne sont actuellement pris en charge que par Excel sur le Web. Toutes les opérations d’enregistrement, de modification et de gestion des scripts sont effectuées sur la plateforme web.

### <a name="script-support-for-excel-on-windows"></a>Prise en charge des scripts pour Excel sur Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="apis"></a>API

Bien que les API JavaScript Office pour les compléments Office et les API Office Scripts partagent certaines fonctionnalités, il s’agit de différentes plateformes. Les API Scripts Office sont un sous-ensemble optimisé et synchrone du modèle d’API JavaScript Excel. La principale différence réside dans l’utilisation du `load`/`sync` paradigme avec les compléments. En outre, les compléments offrent des API pour les événements et un ensemble plus large de fonctionnalités en dehors d’Excel, appelé API communes.

### <a name="events"></a>Events

Les scripts Office ne prennent pas en charge les [événements](/office/dev/add-ins/excel/excel-add-ins-events) au niveau du classeur. Les scripts sont déclenchés par les utilisateurs qui sélectionnent le bouton **Exécuter** pour un script ou via Power Automate. Chaque script exécute le code dans une seule `main` fonction, puis se termine.

### <a name="common-apis"></a>API courantes

Les scripts Office ne peuvent pas utiliser les [API courantes](/javascript/api/office). Si vous avez besoin d’authentification, de fenêtres de dialogue ou d’autres fonctionnalités prises en charge uniquement par les API courantes, vous devrez probablement créer un complément Office au lieu d’un script Office.

## <a name="see-also"></a>Voir aussi

- [Scripts Office dans Excel](../overview/excel.md)
- [Différences entre les scripts Office et les macros VBA](vba-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Créer un complément de volet de tâches Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
