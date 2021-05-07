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
# <a name="how-to-use-macro-files-in-power-automate-flows"></a><span data-ttu-id="62d0a-103">Comment utiliser des fichiers macro dans les flux Power Automate flux</span><span class="sxs-lookup"><span data-stu-id="62d0a-103">How to use macro files in Power Automate flows</span></span>

<span data-ttu-id="62d0a-104">[Power Automate flux](https://flow.microsoft.com/) fournissent [des connecteurs Excel](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) pour vous aider à connecter les fichiers Excel avec le reste de vos données organisationnelles et applications telles que Teams, Outlook et SharePoint.</span><span class="sxs-lookup"><span data-stu-id="62d0a-104">[Power Automate flows](https://flow.microsoft.com/) provide [Excel connectors](https://flow.microsoft.com/connectors/shared_excelonlinebusiness/excel-online-business/) to help connect Excel files with the rest of your organizational data and apps such as Teams, Outlook, and SharePoint.</span></span>

<span data-ttu-id="62d0a-105">Toutefois, les fichiers macro ne peuvent pas être sélectionnés dans la liste finale du fichier (voir un exemple dans la capture d’écran suivante).</span><span class="sxs-lookup"><span data-stu-id="62d0a-105">However, macro files can't be selected in the file dropdown (see an example in the following screenshot).</span></span>

:::image type="content" source="../images/no-xlsm.png" alt-text="L’Power Automate exécuter une action de script indiquant qu’aucun fichier macro n’est sélectionné. L’erreur affichée est « Fichier » est obligatoire.":::

<span data-ttu-id="62d0a-107">Pour contourner ce problème, vous pouvez inclure l’action « Obtenir les métadonnées de fichier » (OneDrive ou SharePoint) et utiliser la propriété ID dans l’action « Exécuter le script », comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="62d0a-107">One way to get around this issue is by including the "Get File Metadata" action (OneDrive or SharePoint) and use the ID property in the "Run Script" action as shown in the following screenshot.</span></span>

:::image type="content" source="../images/xlsm-in-pa.png" alt-text="L’Power Automate exécuter l’action de script montrant le fichier macro sélectionné et aucune erreur de script d’exécuter":::

> [!NOTE]
> <span data-ttu-id="62d0a-109">Certains xlSM (en particulier ceux avec des contrôles ActiveX/formulaire) peuvent ne pas fonctionner dans le connecteur Excel en ligne.</span><span class="sxs-lookup"><span data-stu-id="62d0a-109">Some XLSM (especially the ones with ActiveX/Form controls) may not work in the Excel online connector.</span></span> <span data-ttu-id="62d0a-110">Veillez à tester avant de déployer votre solution.</span><span class="sxs-lookup"><span data-stu-id="62d0a-110">Be sure to test before deploying your solution.</span></span>

## <a name="other-resources"></a><span data-ttu-id="62d0a-111">Autres ressources</span><span class="sxs-lookup"><span data-stu-id="62d0a-111">Other resources</span></span>

<span data-ttu-id="62d0a-112">[Regardez la vidéo YouTube de Sudhi Journalthy sur l’utilisation d’un fichier .xlsm](https://youtu.be/o-H9BbywJQQ)dans une action exécuter un script.</span><span class="sxs-lookup"><span data-stu-id="62d0a-112">[Watch Sudhi Ramamurthy's YouTube video on how use an .xlsm file in a Run Script action](https://youtu.be/o-H9BbywJQQ).</span></span>
