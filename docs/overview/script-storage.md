---
title: Stockage et appartenance des fichiers de scripts Office
description: Informations sur le stockage des scripts Office dans Microsoft OneDrive et leur transfert entre les propriétaires.
ms.date: 11/13/2020
localization_priority: Normal
ms.openlocfilehash: 648f3b2cf7e7d8d3bab2cf07a090e116e267a99a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49346864"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Stockage et appartenance des fichiers de scripts Office

Les scripts Office sont stockés en tant que fichiers **. OSTs** dans votre Microsoft OneDrive. Cela permet à vos scripts d’exister en dehors d’un classeur particulier. Vos paramètres OneDrive contrôlent les accès et les autorisations partagés pour tous les fichiers script **. OSTs** ; indépendamment des paramètres Excel.

## <a name="file-storage"></a>Stockage de fichiers

Les scripts Office sont stockés dans votre espace OneDrive. Les fichiers **. OSTs** se trouvent dans le dossier **scripts/scripts/documents/Office** . Toutes les modifications apportées à ces fichiers **. OSTs** , telles que le changement de nom ou la suppression de fichiers, seront reflétées dans l’éditeur de code et la bibliothèque de scripts.

Les scripts partagés avec l’un de vos classeurs restent dans le OneDrive du créateur de script. Elles ne sont pas copiées dans vos dossiers locaux ou OneDrive lorsque vous exécutez le script partagé dans Excel. Le bouton **créer une copie** de l’éditeur de code enregistre une copie distincte du script dans votre OneDrive. Les modifications apportées à la copie n’affectent pas le script d’origine.

### <a name="script-folders"></a>Dossiers de script

L’ajout de dossiers à OneDrive permet de maintenir l’organisation de vos scripts. Tous les dossiers sous **/documents/Office scripts/** sont affichés sous la section **mes scripts** de l’éditeur de code. Veuillez noter que ces dossiers ne peuvent pas être créés ou supprimés à l’aide de l’éditeur de code. De même, les scripts ne peuvent pas être placés dans des dossiers ou déplacés dans des dossiers à l’aide de l’éditeur de code.

![Certains scripts contenus dans des dossiers, comme affiché dans le volet Office éditeur de code](../images/script-folders.png)

## <a name="file-ownership-and-retention"></a>Propriété et rétention des fichiers

Les scripts Office sont stockés dans le OneDrive d’un utilisateur. Ils suivent les stratégies de rétention et de suppression spécifiées par Microsoft OneDrive. Pour savoir comment gérer les scripts qui ont été créés et partagés par un utilisateur supprimé de votre organisation, consultez [Rétention et suppression de OneDrive](/onedrive/retention-and-deletion).

## <a name="see-also"></a>Voir aussi

- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Annuler les effets d’un script Office](../testing/undo.md)
