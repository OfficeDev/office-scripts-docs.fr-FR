---
title: Office Stockage et propriété des fichiers scripts
description: Informations sur la façon Office scripts sont stockés dans Microsoft OneDrive et transférés entre les propriétaires.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 47b732399c3068bea78b027e01324bbd73a83bc7
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232528"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Stockage et propriété des fichiers scripts

Office Les scripts sont stockés en tant que fichiers **.osts** dans votre Microsoft OneDrive. Cela permet à vos scripts d’exister en dehors d’un workbook particulier. Vos paramètres OneDrive contrôlent l’accès partagé et les autorisations pour tous les fichiers **.osts** de script ; indépendamment de tous Excel paramètres.

## <a name="file-storage"></a>Stockage de fichiers

Vous Office scripts sont stockés dans votre OneDrive. Les **fichiers .osts se** trouvent dans le dossier **/Documents/Office Scripts/.** Toutes les modifications de ces fichiers **.osts,** telles que le changement de nom ou la suppression de fichiers, seront reflétées dans l’éditeur de code et la galerie de scripts.

Les scripts partagés avec l’un de vos workbooks restent dans la OneDrive. Ils ne sont copiés dans aucun de vos dossiers locaux ou OneDrive lorsque vous exécutez le script partagé dans Excel. Le **bouton Effectuer une copie** de l’Éditeur de code enregistre une copie distincte du script dans votre OneDrive. Les modifications apportées à la copie n’affectent pas le script d’origine.

### <a name="script-folders"></a>Dossiers de script

L’ajout de dossiers à OneDrive permet de maintenir l’organisation de vos scripts. Tous les dossiers sous **/Documents/Office Scripts/** sont affichés sous la section **Mes scripts** de l’éditeur de code. Notez que ces dossiers ne peuvent pas être créés ou supprimés à l’aide de l’éditeur de code. De même, les scripts ne peuvent pas être placés dans des dossiers ou déplacés entre dossiers à l’aide de l’Éditeur de code.

:::image type="content" source="../images/script-folders.png" alt-text="Boîte de dialogue Nouveau script dans l’Éditeur de code affichant les scripts contenus dans des dossiers, tel qu’affiché dans le volet Des tâches":::

## <a name="file-ownership-and-retention"></a>Propriété et rétention des fichiers

Office Les scripts sont stockés dans la base de données d’un OneDrive. Ils suivent les stratégies de rétention et de suppression spécifiées par Microsoft OneDrive. Pour savoir comment gérer les scripts qui ont été créés et partagés par un utilisateur supprimé de votre organisation, consultez [Rétention et suppression de OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Voir aussi

- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Annuler les effets d’un script Office](../testing/undo.md)
