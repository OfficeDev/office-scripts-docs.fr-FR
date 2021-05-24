---
title: Office Stockage et propriété des fichiers scripts
description: Informations sur la façon Office scripts sont stockés dans Microsoft OneDrive et transférés entre les propriétaires.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545800"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Stockage et propriété des fichiers scripts

Office Les scripts sont stockés en tant que fichiers **.osts** dans votre Microsoft OneDrive. Ils sont stockés séparément à partir d’un workbook. Pour accorder l’accès à d’autres personnes, [partagez le script avec un Excel de travail.](excel.md#sharing-scripts) Cela signifie que vous liez le script au fichier, et non que vous l’attachez. Toute personne qui a accès au fichier Excel peut également afficher, exécuter ou effectuer une copie du script.

Sauf si vous partagez vos scripts, personne d’autre ne peut y accéder. Vos paramètres OneDrive contrôlent l’accès partagé et les autorisations pour tous les fichiers **.osts** de script, indépendamment des paramètres Excel de script. Les scripts ne peuvent pas être liés à partir d’un disque local ou d’emplacements cloud personnalisés. Office Les scripts reconnaissent et exécutent un script uniquement s’il est dans votre dossier OneDrive ou partagé avec le classeur.

## <a name="file-storage"></a>Stockage de fichiers

Vous Office scripts sont stockés dans votre OneDrive. Les **fichiers .osts se** trouvent dans le dossier **/Documents/Office Scripts/.** Toutes les modifications de ces fichiers **.osts,** telles que le changement de nom ou la suppression de fichiers, seront reflétées dans l’éditeur de code et la galerie de scripts.

Les scripts partagés avec l’un de vos workbooks restent dans la OneDrive. Ils ne sont copiés dans aucun de vos dossiers locaux ou OneDrive lorsque vous exécutez le script partagé dans Excel. Le **bouton Effectuer une copie** de l’Éditeur de code enregistre une copie distincte du script dans votre OneDrive. Les modifications apportées à la copie n’affectent pas le script d’origine.

## <a name="file-ownership-and-retention"></a>Propriété et rétention des fichiers

Office Les scripts sont stockés dans la base de données d’un OneDrive. Ils suivent les stratégies de rétention et de suppression spécifiées par Microsoft OneDrive. Pour savoir comment gérer les scripts qui ont été créés et partagés par un utilisateur supprimé de votre organisation, consultez [Rétention et suppression de OneDrive](/onedrive/retention-and-deletion).

Pendant la modification, les fichiers sont temporairement stockés dans le navigateur. Vous devez enregistrer le script avant de fermer la fenêtre Excel pour l’enregistrer à l’OneDrive emplacement. N’oubliez pas d’enregistrer le fichier après les modifications, sinon ces modifications seront uniquement dans la version du fichier du navigateur.

## <a name="see-also"></a>Voir aussi

- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Annuler les effets des scripts Office scripts](../testing/undo.md)
