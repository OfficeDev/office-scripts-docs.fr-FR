---
title: Scripts Office dans Excel
description: Une brève introduction sur l’enregistreur d’actions et l’éditeur de code pour Office Scripts.
ms.topic: overview
ms.date: 10/05/2022
ms.localizationpriority: high
ms.openlocfilehash: 02a45e5aae468cff2c61e18b35c54ba656d0484b
ms.sourcegitcommit: 64d506257bee282fb01aedbf4d090781b06e4900
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/07/2022
ms.locfileid: "68495473"
---
# <a name="office-scripts-in-excel"></a>Scripts Office dans Excel

Office Scripts dans Excel sur le web vous permet d’automatiser vos tâches quotidiennes. Dans Excel sur le Web, vous pouvez enregistrer vos actions avec l’enregistreur d’actions. Cela crée un script de langage TypeScript qui peut être réexécuté à tout moment. Vous pouvez également créer et modifier des scripts avec l’éditeur de code. Vos scripts peuvent ensuite être partagés au sein de votre organisation afin que vos collègues puissent également automatiser leurs flux de travail.

Cette série de documents vous explique comment utiliser ces outils. Vous allez découvrir comment enregistrer vos actions Excel fréquentes avec l’enregistreur d’actions. Vous découvrirez également comment créer ou mettre à jour vos propres scripts à l’aide de l’éditeur de code.

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

## <a name="requirements"></a>Configuration requise

Pour utiliser les scripts Office, vous devez disposer des éléments suivants.

1. [Excel sur le Web](https://www.office.com/launch/excel) (Excel sur Windows peut uniquement utiliser des scripts Office avec des [boutons de script](../develop/script-buttons.md)).

    > [!TIP]
    > Les scripts Office sont désormais disponibles dans Office sur Windows et sur Mac pour [les insiders Office](https://insider.office.com/).

1. OneDrive Entreprise.
1. Toute licence Microsoft 365 commerciale ou éducative donnant accès aux applications de bureau Microsoft 365, telles que :

    - Applications Microsoft 365 pour les PME
    - Office 365 Business Premium
    - Office 365 ProPlus
    - Office 365 ProPlus pour les appareils
    - Office 365 Entreprise E3
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

> [!TIP]
> L’enregistreur d’actions dispose d’un bouton **Copier en tant que code** pour enregistrer les actions dans le code du script sans enregistrer l’intégralité du script.
>
> :::image type="content" source="../images/action-recorder-copy-code.png" alt-text="Volet des tâches de l’enregistreur d’actions avec le bouton « Copier en tant que code » en surbrillance.":::

Nos [didacticiels](../tutorials/excel-tutorial.md) proposent un apprentissage guidé et structuré des fonctionnalités des scripts Office. Après avoir terminé les didacticiels, lisez [Principes de base des scripts pour les scripts Office dans Excel sur le web](../develop/scripting-fundamentals.md) pour en savoir plus sur l’éditeur de code et sur la rédaction et la modification de vos propres scripts. Pour plus d’informations sur l’Éditeur de code et la manière dont votre code script est interprété, lisez [Environnement Éditeur de code des Scripts Office](code-editor-environment.md).

## <a name="share-office-scripts"></a>Partager des Scripts Office

Les scénarios Office peuvent être partagés avec d'autres utilisateurs d'un classeur Excel. Lorsque vous partagez un script dans un classeur partagé, tous les personnes ayant accès au groupe peuvent également afficher et exécuter votre script. Pour plus d’informations sur le partage et le partage de scripts, voir [Sharing Office Scripts in Excel for the Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b).

:::image type="content" source="../images/script-sharing.png" alt-text="La page des détails du script affichant l'option « Partager avec d'autres dans ce cahier de travail ».":::

Ajoutez des boutons qui exécutent des scripts pour aider vos collègues à découvrir vos solutions précieuses et à les laisser exécuter des scripts dans Excel sur le Bureau. En savoir plus sur les boutons de script dans [Exécuter Scripts avec des boutons](../develop/script-buttons.md).

:::image type="content" source="../images/add-button.png" alt-text="Un bouton de la feuille de calcul qui exécute un script lorsque l’utilisateur clique dessus.":::

> [!NOTE]
> Si vous souhaitez en savoir plus sur le stockage des scripts dans votre espace OneDrive, veuillez consulter la rubrique [Stockage et propriété des fichiers de scripts Office](script-storage.md).

## <a name="connect-office-scripts-to-power-automate"></a>Connecter les scripts Office à Power Automate

[Automatisation de la puissance](https://flow.microsoft.com/) est un service qui vous aide à créer des flux de travail automatisés entre plusieurs applications et services. Les Office Scripts peuvent être utilisés dans ces flux de travail, ce qui vous permet de contrôler vos scénarios en dehors du cahier de travail. Vous pouvez exécuter vos scénarios selon un calendrier, les déclencher en réponse à des courriels, et bien plus encore. Consultez le didacticiel [Exécuter des scripts Office avec Power Automate](../tutorials/excel-power-automate-manual.md) pour découvrir les bases de la connexion de ces services Automation.

## <a name="next-steps"></a>Étapes suivantes

Suivez le [tutoriel sur Office Scripts dans Excel sur le web](../tutorials/excel-tutorial.md) pour découvrir comment créer votre premier script.

## <a name="see-also"></a>Voir aussi

- [Principes de base pour la rédaction de scripts Office en Excel sur le web](../develop/scripting-fundamentals.md)
- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Introduction à Office Scripts dans Excel](https://support.microsoft.com/office/9fbe283d-adb8-4f13-a75b-a81c6baf163a)
- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Centre de développement de scripts Office](https://developer.microsoft.com/office-scripts)
