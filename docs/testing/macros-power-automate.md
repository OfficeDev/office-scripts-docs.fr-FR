---
title: Utiliser des fichiers macro dans les flux Power Automate
description: Découvrez comment utiliser des fichiers macro ou xlsm dans les flux Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: ec1fe00eb9ddc382ae4bc02187de7a36c97288b1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571248"
---
# <a name="how-to-use-macro-files-in-power-automate-flows"></a>Utilisation des fichiers macro dans les flux Power Automate

[Les flux Power Automate](https://flow.microsoft.com/) fournissent [des connecteurs Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) pour vous aider à connecter des fichiers Excel avec le reste de vos données organisationnelles et applications telles que Teams, Outlook et SharePoint.

Toutefois, les fichiers macro ne peuvent pas être sélectionnés dans la liste finale du fichier (voir un exemple dans la capture d’écran suivante).

![Aucune xlsm dans l’action Exécuter le script](../images/no-xlsm.png)

Pour contourner ce problème, vous pouvez inclure l’action « Obtenir les métadonnées de fichier » (OneDrive ou SharePoint) et utiliser la propriété ID dans l’action « Exécuter le script », comme illustré dans la capture d’écran suivante.

![xlsm dans l’action Exécuter le script](../images/xlsm-in-pa.png)

> [!NOTE]
> Certains fichiers XLSM (notamment ceux avec des contrôles ActiveX/formulaire) peuvent ne pas fonctionner dans le connecteur en ligne Excel. Veillez à tester avant de déployer votre solution.

[![Regarder une vidéo sur l’utilisation de XLSM dans l’action Exécuter un script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vidéo sur l’utilisation de XLSM dans l’action Exécuter le script")
