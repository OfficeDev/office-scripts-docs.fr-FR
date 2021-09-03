---
title: Utiliser des fichiers macro dans Power Automate flux
description: Découvrez comment utiliser des fichiers macro ou xlsm dans Power Automate flux.
ms.date: 09/01/2021
localization_priority: Normal
ms.openlocfilehash: 204eb8315f90c0ab0e20a34ec64517fbfbf304b1
ms.sourcegitcommit: 9472e78eb186ceffdaaceb2718d5074ce55a0e74
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2021
ms.locfileid: "58866538"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Comment utiliser des fichiers macro dans les flux Power Automate flux

Le [connecteur Excel Online (Entreprise)](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) dans [Power Automate](https://flow.microsoft.com/) fonctionne généralement uniquement avec les fichiers au format Microsoft Excel feuille de calcul Open XML (.xlsx). Le navigateur de fichiers limite votre sélection .xlsx fichiers à l’intérieur du connecteur. Toutefois, les fichiers macro sont compatibles avec l’action de **script Exécuter** du connecteur si les métadonnées de fichier sont utilisées.

Dans votre flux, utilisez l’action **Obtenir** des métadonnées de fichier à partir des [connecteurs OneDrive Entreprise](https://flow.microsoft.com/connectors/shared_onedriveforbusiness/onedrive-for-business/) ou [SharePoint](https://flow.microsoft.com/connectors/shared_sharepointonline/sharepoint/) fichiers. **L’action de script** Exécuter accepte ces métadonnées en tant que fichier valide. Utilisez le *contenu dynamique de l’ID* renvoyé par l’action  Obtenir les métadonnées du fichier comme argument « Fichier » lors de l’exécution du script. La capture d’écran suivante montre un flux fournissant les métadonnées d’un fichier appelé « Test Macro File.xlsm » vers une action de **script Exécuter.**

:::image type="content" source="../images/xlsm-in-power-automate.png" alt-text="Flux avec une action obtenir des métadonnées de fichier en passant les métadonnées d’un fichier macro à une action de script Exécuter.":::

> [!WARNING]
> Certains fichiers .xlsm, en particulier ceux avec des contrôles ActiveX ou formulaire, peuvent ne pas fonctionner dans le connecteur Excel en ligne. Veillez à tester avant de déployer votre solution.

## <a name="other-resources"></a>Autres ressources

[Regardez la vidéo YouTube de Sudhi Journalthy sur l’utilisation d’un fichier .xlsm](https://youtu.be/o-H9BbywJQQ)dans une action exécuter un script.
