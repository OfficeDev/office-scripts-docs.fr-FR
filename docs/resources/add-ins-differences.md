---
title: Différences entre les scripts Office et les compléments Office
description: Différences de comportement et d'API entre les scripts Office et les add-ins Office.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 96af98ca9f247406c5cc916f38892c318d33c560
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755097"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Différences entre les scripts Office et les compléments Office

Les add-ins Office et les scripts Office ont beaucoup en commun. Ils offrent tous deux un contrôle automatisé d'un workbook Excel sur une API JavaScript. Toutefois, les API Office Scripts sont une version spécialisée et synchrone de l'API JavaScript Pour Office.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les zones de focus pour différentes solutions d'extensibilité Office. Les scripts Office et les applications web Office sont axés sur le web et la collaboration, mais les scripts Office sont pris en charge par les utilisateurs finaux (tandis que les applications web Office ciblent les développeurs professionnels).":::

Les scripts Office s'exécutent jusqu'à la fin avec un bouton manuel ou à l'étape de [Power Automate,](https://flow.microsoft.com/)tandis que les compl?ments Office persistent pendant que leurs volets office sont ouverts. Cela signifie que les add-ins peuvent maintenir l'état pendant une session, tandis que les scripts Office ne conservent pas un état interne entre les séquences. Si vous constatez que votre extension Excel doit dépasser les fonctionnalités de la plateforme d'écriture de scripts, consultez la documentation des [add-ins Office](/office/dev/add-ins) pour en savoir plus sur les extensions Office.

Le reste de cet article décrit les principales différences entre les add-ins Office et les scripts Office.

## <a name="platform-support"></a>Prise en charge de la plateforme

Les add-ins Office sont sur plusieurs plateformes. Ils fonctionnent sur les plateformes windows de bureau, Mac, iOS et web et offrent la même expérience sur chacune d'elles. Toute exception à cette règle est notée dans la documentation de l'API individuelle.

Les scripts Office sont actuellement uniquement pris en charge par Excel sur le web. L'enregistrement, la modification et l'exécution s'exécutent sur la plateforme web.

## <a name="apis"></a>API

Il n'existe aucune version synchrone des API JavaScript pour Office pour les add-ins Office. Les API Standard Office Scripts sont propres à la plateforme et ont de nombreuses optimisations et modifications pour éviter l'utilisation du `load` / `sync` paradigme.

Certaines API [JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true) sont compatibles avec les API Async des [scripts Office.](../develop/excel-async-model.md) Certains exemples et blocs de code de add-in peuvent être portés vers des `Excel.run` blocs avec une traduction minimale. Bien que les deux plateformes partagent des fonctionnalités, il existe des lacunes. Les deux principaux ensembles d'API que les add-ins Office ont, mais les scripts Office ne sont pas des événements et les API communes.

### <a name="events"></a>Événements

Les scripts Office ne supportent pas les [événements.](/office/dev/add-ins/excel/excel-add-ins-events) Chaque script exécute le code dans une seule `main` méthode, puis se termine. Elle ne se réactive pas lorsque des événements sont déclenchés et, par conséquent, ne peut pas enregistrer d'événements.

### <a name="common-apis"></a>API courantes

Les scripts Office ne peuvent pas utiliser [les API communes.](/javascript/api/office) Si vous avez besoin d'une authentification, de fenêtres de boîte de dialogue ou d'autres fonctionnalités uniquement pris en charge par les API communes, vous devrez probablement créer un add-in Office au lieu d'un script Office.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Différences entre les scripts Office et les macros VBA](vba-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Créer un complément de volet de tâches Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
