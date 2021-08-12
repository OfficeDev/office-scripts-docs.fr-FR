---
title: Office Stockage et propriété des fichiers scripts
description: Informations sur la façon Office scripts sont stockés dans Microsoft OneDrive et transférés entre les propriétaires.
ms.date: 06/04/2021
localization_priority: Normal
ms.openlocfilehash: 6343b5bad366d9e4c4f349622a33b062de9c8ddd7877c3d40a49635d6aaef9cf
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847295"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Stockage et propriété des fichiers scripts

Office Les scripts sont stockés en tant que fichiers **.osts** dans votre Microsoft OneDrive. Ils sont stockés séparément à partir d’un workbook. Pour accorder l’accès à d’autres personnes, [partagez le script avec un Excel de travail.](excel.md#sharing-scripts) Cela signifie que vous liez le script au fichier, et non que vous l’attachez. Toute personne qui a accès au fichier Excel peut également afficher, exécuter ou effectuer une copie du script.

Sauf si vous partagez vos scripts, personne d’autre ne peut y accéder. Vos OneDrive contrôlent l’accès partagé et les autorisations pour tous les fichiers **.osts** de script, indépendamment des paramètres Excel de script. Les scripts ne peuvent pas être liés à partir d’un disque local ou d’emplacements cloud personnalisés. Office Les scripts reconnaissent et exécutent un script uniquement s’il est dans votre dossier OneDrive ou partagé avec le classeur.

## <a name="file-storage"></a>Stockage de fichiers

Vous Office scripts sont stockés dans votre OneDrive. Les **fichiers .osts se** trouvent dans le dossier **/Documents/Office Scripts/.** Toutes les modifications de ces fichiers **.osts,** telles que le changement de nom ou la suppression de fichiers, seront reflétées dans l’éditeur de code et la galerie de scripts.

Les scripts partagés avec l’un de vos workbooks restent dans la OneDrive. Ils ne sont copiés dans aucun de vos dossiers locaux ou OneDrive lorsque vous exécutez le script partagé dans Excel. Le **bouton Effectuer une copie** de l’Éditeur de code enregistre une copie distincte du script dans votre OneDrive. Les modifications apportées à la copie n’affectent pas le script d’origine.

### <a name="restore-deleted-scripts"></a>Restaurer des scripts supprimés

Lorsque vous supprimez un script dans Excel, il est OneDrive corbeille. Pour restaurer un script supprimé, suivez les étapes répertoriées dans Restaurer les fichiers ou [dossiers supprimés dans OneDrive](https://support.microsoft.com/office/restore-deleted-files-or-folders-in-onedrive-949ada80-0026-4db3-a953-c99083e6a84f). La restauration **d’un fichier .osts** le renvoie à la liste **Tous les scripts.**

Un script supprimé n’est pas partagé avec le workbook. Lorsque vous restituer un script, il ne **conserve pas** son accès au script. Vous devrez partager à nouveau le script.

Les scripts restaurés fonctionnent toujours comme prévu avec Power Automate flux. Vous n’avez pas besoin de recréer le connecteur de flux.

## <a name="file-ownership-and-retention"></a>Propriété et rétention des fichiers

Office Les scripts sont stockés dans la base de données d’un OneDrive. Ils suivent les stratégies de rétention et de suppression spécifiées par Microsoft OneDrive. Pour savoir comment gérer les scripts qui ont été créés et partagés par un utilisateur supprimé de votre organisation, consultez [Rétention et suppression de OneDrive](/onedrive/retention-and-deletion).

Pendant la modification, les fichiers sont temporairement stockés dans le navigateur. Vous devez enregistrer le script avant de fermer la fenêtre Excel pour l’enregistrer à l’OneDrive emplacement. N’oubliez pas d’enregistrer le fichier après les modifications, sinon ces modifications seront uniquement dans la version du fichier du navigateur.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditer Office’utilisation des scripts au niveau de l’administrateur

Découvrez les locataires qui utilisent Office scripts avec le journal d’audit dans le centre de conformité. Pour découvrir comment utiliser cet outil, consultez le journal d’audit dans le Centre de [sécurité & conformité.](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log)

Pour rechercher les personnes qui utilisent Office scripts avec l’outil de recherche, ajoutez le champ Fichier, Dossier `.osts` **ou Site.** Cette opération recherche tous les fichiers avec l’extension Office Scripts. Si une personne de votre organisation a utilisé la fonctionnalité Office Scripts, l’activité de l’utilisateur s’affiche dans les résultats de recherche du journal d’audit.

> [!NOTE]
> L’exécution d’un script n’est actuellement pas enregistrée. Seules les actions créer, afficher et modifier sont enregistrées.

## <a name="see-also"></a>Voir aussi

- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Annuler les effets des scripts Office](../testing/undo.md)
