---
title: Office Scripts dans Excel sur le web
description: Une brève introduction sur l’enregistreur d’actions et l’éditeur de code pour Office Scripts.
ms.date: 07/04/2021
localization_priority: Priority
ms.openlocfilehash: 30058bb2e372e4ad0e344de6202462a744f4e65df9964a8b2329f8b3e0cc27f1
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847076"
---
# <a name="office-scripts-in-excel-on-the-web"></a>Office Scripts dans Excel sur le web

Office Scripts dans Excel sur le web vous permet d’automatiser vos tâches quotidiennes. Vous pouvez enregistrer les actions Excel avec l’enregistreur d’actions, ce qui crée un script linguistique TypeScript. Vous pouvez également créer et modifier des scripts avec l’éditeur de code. Vos scripts peuvent ensuite être partagés au sein de votre organisation afin que vos collègues puissent également automatiser leurs flux de travail.

Cette série de documents vous explique comment utiliser ces outils. Vous allez découvrir comment enregistrer vos actions Excel fréquentes avec l’enregistreur d’actions. Vous découvrirez également comment créer ou mettre à jour vos propres scripts à l’aide de l’éditeur de code.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a>Configuration requise

Pour utiliser les scripts Office, vous devez disposer des éléments suivants.

1. [Excel sur le web](https://www.office.com/launch/excel) (les autres plateformes, telles que le bureau, ne sont pas prises en charge).
1. OneDrive Entreprise.
1. Toute licence Microsoft 365 commerciale ou éducative donnant accès aux applications de bureau Microsoft 365, telles que :

    - Applications Microsoft 365 pour les PME
    - Office 365 Business Premium
    - Office 365 ProPlus
    - Office 365 ProPlus pour les appareils
    - Office 365 Entreprise E3
    - Office 365 Entreprise E5
    - Office 365 A3
    - Office 365 A5

> [!NOTE]
> Si vous répondez à ces critères et que l’onglet **Automatiser** n’apparaît pas, il est possible que votre administrateur ait désactivé la fonctionnalité ou qu’un autre problème se soit produit dans votre environnement. Veuillez suivre les étapes décrites dans [L’onglet Automatiser n’apparaît pas ou les scripts Office ne sont pas disponibles](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable) pour commencer à utiliser les scripts Office.

## <a name="when-to-use-office-scripts"></a>Quand utiliser les scripts Office

Les scripts permettent d’enregistrer et de reproduire les actions Excel sur des classeurs et des feuilles de calcul différents. Si votre travail comporte des tâches répétitives, vous pouvez les automatiser facilement à l’aide d’un script Office. Exécutez votre script en appuyant sur un bouton dans Excel ou combinez-le avec Power Automate pour rationaliser l'ensemble de votre flux de travail.

Par exemple, vous débutez votre journée de travail en ouvrant un fichier .csv à partir d’un site de comptabilité dans Excel. Vous consacrez ensuite plusieurs minutes à supprimer les colonnes superflues, à mettre en forme un tableau, à ajouter des formules et à créer un tableau croisé dynamique dans une nouvelle feuille de calcul. Ces actions que vous répétez quotidiennement peuvent être enregistrées une seule fois à l’aide de l’enregistreur d’actions. Dès lors, l’exécution du script se chargera de l’intégralité de la conversion .csv. Non seulement, vous supprimez le risque d’oublier les étapes, mais vous pouvez aussi partager le procédé avec d’autres personnes sans avoir à leur apprendre ce dernier. Les scripts Office vous permettent d’automatiser vos tâches courantes afin que vous et votre équipe puissiez être plus efficaces et productifs.

## <a name="action-recorder"></a>Enregistreur d’actions

:::image type="content" source="../images/action-recorder-intro.png" alt-text="Liste d’actions enregistrées par l’enregistreur d’action.":::

La fonctionnalité Enregistreur d’actions vous permet d’enregistrer les actions que vous effectuez dans Excel sous forme de script. Quand l’enregistreur d’actions est en cours d’exécution, vous pouvez capturer les actions Excel effectuées lorsque vous modifiez des cellules, la mise en forme et créez des tableaux. Le script obtenu peut être exécuté sur d’autres feuilles de calcul et classeurs pour recréer vos actions d’origine.

## <a name="code-editor"></a>Éditeur de code

:::image type="content" source="../images/code-editor-intro.png" alt-text="Éditeur de code affichant le code du script utilisé dans ce didacticiel.":::

Tous les scripts enregistrés avec l’enregistreur d’actions peuvent être modifiés via l’éditeur de code. Cela vous permet d’affiner et de personnaliser le script pour l’adapter à vos besoins. Vous pouvez également ajouter une logique et des fonctionnalités qui ne sont pas directement accessibles via l’interface utilisateur d’Excel, comme les instructions conditionnelles (si/sinon) et les boucles.

Un moyen simple de commencer à apprendre les fonctionnalités de Office Scripts consiste à enregistrer les scripts dans Excel sur le web et à afficher le code obtenu. Vous pouvez également suivre nos [tutoriels](../tutorials/excel-tutorial.md) pour découvrir des instructions qui vous guideront de manière plus structurée. 

Après avoir terminé le didacticiel, lisez [Principes de base de l’écriture de scripts Office dans Excel sur le web](../develop/scripting-fundamentals.md) pour en savoir plus sur l’Éditeur de code et la rédaction et la modification de vos propres scripts. Pour plus d’informations sur l’Éditeur de code et la manière dont votre code script est interprété, lisez [Environnement Éditeur de code des Scripts Office](code-editor-environment.md).

## <a name="sharing-scripts"></a>Partage de scénarios

:::image type="content" source="../images/script-sharing.png" alt-text="La page des détails du scénario montrant l'option « Partager avec d'autres dans ce cahier de travail ».":::

Les scénarios Office peuvent être partagés avec d'autres utilisateurs d'un classeur Excel. Lorsque vous partagez un script dans un classeur partagé, tous les personnes ayant accès au groupe peuvent également afficher et exécuter votre script.

Si vous souhaitez en savoir plus sur le partage et le non-partage de scripts, veuillez consulter l'article [Partage de scripts Office dans Excel pour le web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b).

> [!NOTE]
> Si vous souhaitez en savoir plus sur le stockage des scripts dans votre espace OneDrive, veuillez consulter la rubrique [Stockage et propriété des fichiers de scripts Office](script-storage.md).

## <a name="connecting-office-scripts-to-power-automate"></a>Connecter les scripts Office à Power Automate

[Automatisation de la puissance](https://flow.microsoft.com/) est un service qui vous aide à créer des flux de travail automatisés entre plusieurs applications et services. Les Office Scripts peuvent être utilisés dans ces flux de travail, ce qui vous permet de contrôler vos scénarios en dehors du cahier de travail. Vous pouvez exécuter vos scénarios selon un calendrier, les déclencher en réponse à des courriels, et bien plus encore. Visitez le [Exécutez des scénarios Office en Excel sur le web avec Power Automate](../tutorials/excel-power-automate-manual.md) pour apprendre les bases de la connexion de ces services d'automatisation.

## <a name="next-steps"></a>Étapes suivantes

Suivez le [tutoriel sur Office Scripts dans Excel sur le web](../tutorials/excel-tutorial.md) pour découvrir comment créer votre premier script.

## <a name="see-also"></a>Voir aussi

- [Principes de base pour la rédaction de scripts Office en Excel sur le web](../develop/scripting-fundamentals.md)
- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Introduction à Office Scripts dans Excel (sur support.office.com)](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Centre de développement de scripts Office](https://developer.microsoft.com/office-scripts)

