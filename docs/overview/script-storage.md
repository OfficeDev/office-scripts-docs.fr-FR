---
title: Office Stockage et propriété de fichiers scripts
description: Informations sur la façon dont Office scripts sont stockés dans Microsoft OneDrive et transférés entre les propriétaires.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 556d784dc1fe64873866c49ab2726a4c68abc1a7
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545800"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Stockage et propriété de fichiers scripts

Office Les scripts sont stockés **sous forme de fichiers .osts** dans Microsoft OneDrive. Ils sont stockés séparément d’un cahier de travail. Pour donner accès aux autres, [partagez le script avec un Excel de travail](excel.md#sharing-scripts). Cela signifie que vous reliez le script au fichier, et non l’attache. Quiconque a accès au fichier Excel sera également en mesure de visualiser, exécuter ou faire une copie du script.

Sauf si vous partagez vos scripts, personne d’autre ne peut y accéder. Vos OneDrive contrôlent l’accès partagé et les autorisations pour tous les fichiers script **.osts,** indépendamment de tous les paramètres Excel. Les scripts ne peuvent pas être liés à partir d’un disque local ou d’emplacements cloud personnalisés. Office Les scripts ne reconnaissent et exécutent un script que s’il est dans votre dossier OneDrive ou partagé avec le cahier de travail.

## <a name="file-storage"></a>Stockage de fichiers

Vous Office scripts sont stockés dans votre OneDrive. Les **fichiers .osts** se trouvent dans le **fichier /Documents/Office Scripts/dossier.** Toutes les modifications apportées à ces **fichiers .osts,** telles que le changement de nom ou la suppression de fichiers, seront reflétées dans l’éditeur de code et la galerie de scripts.

Les scripts qui sont partagés avec l’un de vos cahiers restent dans le processus de OneDrive. Ils ne sont copiés sur aucun de vos dossiers locaux ou OneDrive lorsque vous exécutez le script partagé dans Excel. Le **bouton Faire une** copie de l’éditeur de code enregistre une copie séparée du script dans votre OneDrive. Les modifications apportées à la copie n’affectent pas le script d’origine.

## <a name="file-ownership-and-retention"></a>Propriété et conservation des fichiers

Office Les scripts sont stockés dans les données d’un OneDrive. Ils suivent les politiques de rétention et de suppression spécifiées par Microsoft OneDrive. Pour savoir comment gérer les scripts qui ont été créés et partagés par un utilisateur supprimé de votre organisation, consultez [Rétention et suppression de OneDrive](/onedrive/retention-and-deletion).

Pendant l’édition, les fichiers sont stockés temporairement dans le navigateur. Vous devez enregistrer le script avant de fermer la fenêtre Excel pour l’enregistrer à l’emplacement OneDrive’écriture. N’oubliez pas d’enregistrer le fichier après les modifications, sinon ces modifications ne seront que dans la version du navigateur du fichier.

## <a name="see-also"></a>Voir aussi

- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/sharing-office-scripts-in-excel-for-the-web-226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [Annuler les effets des scripts Office texte](../testing/undo.md)
