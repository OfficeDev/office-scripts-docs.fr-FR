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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="c3116-103">Utilisation des fichiers macro dans les flux Power Automate</span><span class="sxs-lookup"><span data-stu-id="c3116-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="c3116-104">[Les flux Power Automate](https://flow.microsoft.com/) fournissent [des connecteurs Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) pour vous aider à connecter des fichiers Excel avec le reste de vos données organisationnelles et applications telles que Teams, Outlook et SharePoint.</span><span class="sxs-lookup"><span data-stu-id="c3116-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="c3116-105">Toutefois, les fichiers macro ne peuvent pas être sélectionnés dans la liste finale du fichier (voir un exemple dans la capture d’écran suivante).</span><span class="sxs-lookup"><span data-stu-id="c3116-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

![Aucune xlsm dans l’action Exécuter le script](../images/no-xlsm.png)

<span data-ttu-id="c3116-107">Pour contourner ce problème, vous pouvez inclure l’action « Obtenir les métadonnées de fichier » (OneDrive ou SharePoint) et utiliser la propriété ID dans l’action « Exécuter le script », comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="c3116-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

![xlsm dans l’action Exécuter le script](../images/xlsm-in-pa.png)

> [!NOTE]
> <span data-ttu-id="c3116-109">Certains fichiers XLSM (notamment ceux avec des contrôles ActiveX/formulaire) peuvent ne pas fonctionner dans le connecteur en ligne Excel.</span><span class="sxs-lookup"><span data-stu-id="c3116-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="c3116-110">Veillez à tester avant de déployer votre solution.</span><span class="sxs-lookup"><span data-stu-id="c3116-110">Be sure to test before deploying your solution.</span></span>

<span data-ttu-id="c3116-111">[![Regarder une vidéo sur l’utilisation de XLSM dans l’action Exécuter un script](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Vidéo sur l’utilisation de XLSM dans l’action Exécuter le script")</span><span class="sxs-lookup"><span data-stu-id="c3116-111">[![Watch video about using XLSM in Run Script action](../images/xlsm-vid.png)](https://youtu.be/o-H9BbywJQQ "Video about using XLSM in Run Script action")</span></span>
