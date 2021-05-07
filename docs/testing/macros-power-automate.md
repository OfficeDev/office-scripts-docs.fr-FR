---
title: Utiliser des fichiers macro dans Power Automate flux
description: Découvrez comment utiliser des fichiers macro ou xlsm dans Power Automate flux.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b232a1d31a7ff6e28016c5e28fd8a83c8d3f1859
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232654"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Comment utiliser des fichiers macro dans les flux Power Automate flux

[Power Automate flux](https://flow.microsoft.com/) fournissent [des connecteurs Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) pour vous aider à connecter les fichiers Excel avec le reste de vos données organisationnelles et applications telles que Teams, Outlook et SharePoint.

Toutefois, les fichiers macro ne peuvent pas être sélectionnés dans la liste finale du fichier (voir un exemple dans la capture d’écran suivante).

:::image type="content" source="../images/no-xlsm.png" alt-text="L’Power Automate exécuter une action de script indiquant qu’aucun fichier macro n’est sélectionné. L’erreur affichée est « Fichier » est obligatoire.":::

Pour contourner ce problème, vous pouvez inclure l’action « Obtenir les métadonnées de fichier » (OneDrive ou SharePoint) et utiliser la propriété ID dans l’action « Exécuter le script », comme illustré dans la capture d’écran suivante.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="L’Power Automate exécuter l’action de script montrant le fichier macro sélectionné et aucune erreur de script d’exécuter":::

> [!NOTE]
> Certains xlSM (en particulier ceux avec des contrôles ActiveX/formulaire) peuvent ne pas fonctionner dans le connecteur Excel en ligne. Veillez à tester avant de déployer votre solution.

## <a name="other-resources"></a>Autres ressources

[Regardez la vidéo YouTube de Sudhi Journalthy sur l’utilisation d’un fichier .xlsm](https://youtu.be/o-H9BbywJQQ)dans une action exécuter un script.
