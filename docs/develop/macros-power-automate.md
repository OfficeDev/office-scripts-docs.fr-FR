---
title: Utiliser des fichiers activés pour les macros dans Power Automate flux
description: Découvrez comment utiliser des fichiers activés pour les macros, ou des fichiers .xlsm, Power Automate flux.
ms.date: 03/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f2ecefe9fb97d1c5514ddb52c3cbcd0596df426
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585743"
---
# <a name="how-to-use-macro-enabled-files-in-power-automate-flows"></a>Utilisation de fichiers activés pour les macros dans Power Automate flux

Vous pouvez intégrer vos fichiers .xlsm à un flux Power Automate de données. Cela vous permet de commencer à convertir vos solutions d’automatisation existantes en formats web. Notez que les macros contenues dans les fichiers .xslm ne peuvent pas être Power Automate. Seuls Office scripts sont activés ici.

Le [connecteur Excel Online (Entreprise)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) dans [Power Automate](https://flow.microsoft.com/) est généralement limité aux fichiers au format Microsoft Excel feuille de calcul Open XML (.xlsx). Son navigateur de fichiers vous permet uniquement de sélectionner .xlsx fichiers. Toutefois, les fichiers compatibles avec les macros sont compatibles avec l’action de **script** Exécuter du connecteur si les métadonnées de fichier sont utilisées.

Dans votre flux, utilisez l’action **Obtenir** les métadonnées de fichier à partir des [connecteurs OneDrive Entreprise](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) ou [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) fichiers. **L’action de script** Exécuter accepte ces métadonnées en tant que fichier valide. Utilisez le *contenu dynamique de l’ID* renvoyé par **l’action Obtenir** les métadonnées du fichier comme argument « Fichier » lors de l’exécution du script. La capture d’écran suivante montre un flux fournissant les métadonnées d’un fichier appelé « Test Macro File.xlsm » vers une action **de script Exécuter** .

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Flux avec une action obtenir des métadonnées de fichier en passant les métadonnées d’un fichier macro à une action de script Exécuter.":::

> [!WARNING]
> Certains fichiers .xlsm, en particulier ceux avec des contrôles ActiveX formulaire, peuvent ne pas fonctionner dans le connecteur Excel en ligne. Veillez à tester avant de déployer votre solution.

## <a name="other-resources"></a>Autres ressources

[Regardez la vidéo YouTube de Sudhi Journaly sur l’utilisation d’un fichier .xlsm dans une action Exécuter un script](https://youtu.be/o-H9BbywJQQ).
