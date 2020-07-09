---
title: Office Scripts dans Excel sur le web
description: Une brève introduction sur l’enregistreur d’actions et l’éditeur de code pour Office Scripts.
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: 046dd4eac0cce14117da75199841f0b2f72031bc
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043404"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a>Office Scripts dans Excel sur le web (Aperçu)

Office Scripts dans Excel sur le web vous permet d’automatiser vos tâches quotidiennes. Vous pouvez enregistrer les actions Excel avec l’enregistreur d’actions, ce qui crée un script. Vous pouvez également créer et modifier des scripts avec l’éditeur de code. Vos scripts peuvent ensuite être partagés au sein de votre organisation afin que vos collègues puissent également automatiser leurs flux de travail.

Cette série de documents vous explique comment utiliser ces outils. Vous allez découvrir comment enregistrer vos actions Excel fréquentes avec l’enregistreur d’actions. Vous découvrirez également comment créer ou mettre à jour vos propres scripts à l’aide de l’éditeur de code.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="when-to-use-office-scripts"></a>Quand utiliser Office Scripts

Les scripts permettent d’enregistrer et de reproduire les actions Excel sur des classeurs et des feuilles de calcul différents. Si vous avez l’impression de faire tout le temps les mêmes choses, un script Office peut vous aider à réduire l’ensemble de votre workflow grâce à un seul bouton.

Par exemple, vous débutez votre journée de travail en ouvrant un fichier .csv à partir d’un site de comptabilité dans Excel. Vous consacrez ensuite plusieurs minutes à supprimer les colonnes superflues, à mettre en forme un tableau, à ajouter des formules et à créer un tableau croisé dynamique dans une nouvelle feuille de calcul. Ces actions que vous répétez quotidiennement peuvent être enregistrées une seule fois à l’aide de l’enregistreur d’actions. Dès lors, l’exécution du script se chargera de l’intégralité de la conversion .csv. Non seulement, vous supprimez le risque d’oublier les étapes, mais vous pouvez aussi partager le procédé avec d’autres personnes sans avoir à leur apprendre ce dernier. Les scripts Office automatisent vos tâches courantes afin que vous et votre équipe puissiez être plus efficaces et productifs.

## <a name="action-recorder"></a>Enregistreur d’actions

![L’enregistreur d’actions après avoir enregistré plusieurs actions.](../images/action-recorder-intro.png)

L’enregistreur d’actions enregistre les actions que vous effectuez dans Excel et les traduit dans un script. Quand l’enregistreur d’actions est en cours d’exécution, vous pouvez capturer les actions Excel effectuées lorsque vous modifiez des cellules, la mise en forme et créez des tableaux. Le script obtenu peut être exécuté sur d’autres feuilles de calcul et classeurs pour recréer vos actions d’origine.

## <a name="code-editor"></a>Éditeur de code

![L’éditeur de code affichant le code du script ci-dessus.](../images/code-editor-intro.png)

Tous les scripts enregistrés avec l’enregistreur d’actions peuvent être modifiés via l’éditeur de code. Cela vous permet d’affiner et de personnaliser le script pour l’adapter à vos besoins. Vous pouvez également ajouter une logique et des fonctionnalités qui ne sont pas directement accessibles via l’interface utilisateur d’Excel, comme les instructions conditionnelles (si/sinon) et les boucles.

Un moyen simple de commencer à apprendre les fonctionnalités de Office Scripts consiste à enregistrer les scripts dans Excel sur le web et à afficher le code obtenu. Vous pouvez également suivre nos [tutoriels](../tutorials/excel-tutorial.md) pour découvrir des instructions qui vous guideront de manière plus structurée. 

## <a name="sharing-scripts"></a>Partage de scénarios

![La page des détails du scénario montrant l'option «Partager avec d'autres dans ce cahier de travail».](../images/script-sharing.png)

Les scénarios Office peuvent être partagés avec d'autres utilisateurs d'un classeur Excel. Lorsque vous partagez un scénario avec d'autres personnes dans un cahier de travail, le scénario est joint au cahier. Vos scénarios sont stockés dans votre OneDrive, et lorsque vous en partagez un, vous créez un lien vers celui-ci dans le cahier de travail que vous avez ouvert.

Plus de détails sur le partage et le non-partage de scripts peuvent être trouvés dans l'article [Partager des scripts de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US).

## <a name="connecting-office-scripts-to-power-automate"></a>Connecter les scripts de bureau à l'automatisation de la puissance

[Automatisation de la puissance](https://flow.microsoft.com/) est un service qui vous aide à créer des flux de travail automatisés entre plusieurs applications et services. Les Office Scripts peuvent être utilisés dans ces flux de travail, ce qui vous permet de contrôler vos scénarios en dehors du cahier de travail. Vous pouvez exécuter vos scénarios selon un calendrier, les déclencher en réponse à des courriels, et bien plus encore. Visitez le [Exécutez des scénarios Office en Excel sur le web avec Power Automate](../tutorials/excel-power-automate-manual.md) pour apprendre les bases de la connexion de ces services d'automatisation.

## <a name="next-steps"></a>Étapes suivantes

Suivez le [tutoriel sur Office Scripts dans Excel sur le web](../tutorials/excel-tutorial.md) pour découvrir comment créer vos premiers scripts Office.

## <a name="see-also"></a>Voir aussi

- [Principes de base pour la rédaction de scripts Office en Excel sur le web](../develop/scripting-fundamentals.md)
- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Introduction à Office Scripts dans Excel (sur support.office.com)](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b?storagetype=live&ui=en-US&rs=en-US&ad=US)
