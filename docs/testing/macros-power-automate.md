---
title: Utiliser des fichiers macro dans les flux Power Automate
description: Découvrez comment utiliser des fichiers macro ou xlsm dans les flux Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: a7929fc485ae2118d30a4f2783538d0e04deca2a
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755013"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Utilisation des fichiers macro dans les flux Power Automate

[Les flux Power Automate](https://flow.microsoft.com/) fournissent [des connecteurs Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) pour vous aider à connecter des fichiers Excel avec le reste de vos données organisationnelles et applications telles que Teams, Outlook et SharePoint.

Toutefois, les fichiers macro ne peuvent pas être sélectionnés dans la liste finale du fichier (voir un exemple dans la capture d'écran suivante).

:::image type="content" source="../images/no-xlsm.png" alt-text="Action de script Exécuter Power Automate indiquant qu'aucun fichier macro n'est sélectionné. L'erreur affichée est « Fichier » est obligatoire.":::

Pour contourner ce problème, vous pouvez inclure l'action « Obtenir les métadonnées de fichier » (OneDrive ou SharePoint) et utiliser la propriété ID dans l'action « Exécuter le script », comme illustré dans la capture d'écran suivante.

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="Action de script Exécuter Power Automate montrant le fichier macro sélectionné et aucune erreur de script Exécuter.":::

> [!NOTE]
> Certains fichiers XLSM (notamment ceux avec des contrôles ActiveX/formulaire) peuvent ne pas fonctionner dans le connecteur en ligne Excel. Veillez à tester avant de déployer votre solution.

[![Regarder une vidéo sur l'utilisation de XLSM dans l'action Exécuter un script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vidéo sur l'utilisation de XLSM dans l'action Exécuter le script")
