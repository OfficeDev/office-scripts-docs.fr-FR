---
title: Office Le stockage et la propriété des fichiers scripts
description: Informations sur la façon dont les scripts Office sont stockés dans Microsoft OneDrive et transférés entre propriétaires.
ms.date: 05/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e2bc89db54ee5520c3b911ebd0f182777a78e2b
ms.sourcegitcommit: 8ae932e8b4e521fec8576ab16126eb9fe22a8dd7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2022
ms.locfileid: "65310756"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Office Le stockage et la propriété des fichiers scripts

Office scripts sont stockés sous forme de fichiers **.osts** dans votre Microsoft OneDrive. Ils sont stockés séparément d’un classeur. Pour permettre à d’autres utilisateurs [d’accéder, partagez le script avec un classeur Excel](excel.md#share-office-scripts). Cela signifie que vous liez le script avec le fichier, et non l’attachez. Quiconque a accès au fichier Excel sera également en mesure d’afficher, d’exécuter ou d’effectuer une copie du script.

Sauf si vous partagez vos scripts, personne d’autre ne peut y accéder. Vos paramètres OneDrive contrôlent l’accès partagé et les autorisations pour tous les fichiers **.osts** de script, indépendamment des paramètres Excel. Les scripts ne peuvent pas être liés à partir d’un disque local ou d’emplacements cloud personnalisés. Office Scripts reconnaît et exécute un script uniquement s’il se trouve dans votre dossier OneDrive ou qu’il est partagé avec le classeur.

## <a name="file-storage"></a>Stockage de fichiers

Vous Office scripts sont stockés dans votre OneDrive. Les fichiers **.osts** se trouvent dans le dossier **/Documents/Office Scripts/**. Toutes les modifications apportées à ces fichiers **.osts** , telles que le changement de nom ou la suppression de fichiers, sont répercutées dans l’Éditeur de code et la galerie de scripts.

Les scripts partagés avec l’un de vos classeurs restent dans le OneDrive du créateur du script. Ils ne sont copiés dans aucun de vos dossiers locaux ou OneDrive lorsque vous exécutez le script partagé dans Excel. Le bouton **Créer une copie** de l’éditeur de code enregistre une copie distincte du script dans votre OneDrive. Les modifications apportées à la copie n’affectent pas le script d’origine.

### <a name="restore-deleted-scripts"></a>Restaurer des scripts supprimés

Lorsque vous supprimez un script dans Excel, il est envoyé à votre corbeille OneDrive. Pour restaurer un script supprimé, suivez les étapes répertoriées dans [Restaurer les fichiers ou dossiers supprimés dans OneDrive](https://support.microsoft.com/office/949ada80-0026-4db3-a953-c99083e6a84f). La restauration d’un fichier **.osts** le renvoie à la liste **Tous les scripts** .

Un script supprimé n’est pas partagé avec le classeur. Lorsque vous restaurez un script, il ne conserve **pas** son accès au script. Vous devrez partager à nouveau le script.

Les scripts restaurés fonctionnent toujours comme prévu avec Power Automate flux. Vous n’avez pas besoin de recréer le connecteur de flux.

## <a name="file-ownership-and-retention"></a>Propriété et rétention des fichiers

Office scripts sont stockés dans le OneDrive d’un utilisateur. Ils suivent les stratégies de rétention et de suppression spécifiées par Microsoft OneDrive. Pour savoir comment gérer les scripts qui ont été créés et partagés par un utilisateur supprimé de votre organisation, consultez [Rétention et suppression de OneDrive](/onedrive/retention-and-deletion).

Pendant la modification, les fichiers sont temporairement stockés dans le navigateur. Vous devez enregistrer le script avant de fermer la fenêtre Excel pour l’enregistrer à l’emplacement OneDrive. N’oubliez pas d’enregistrer le fichier après les modifications, sinon ces modifications se trouveront uniquement dans la version du fichier du navigateur.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditer l’utilisation des scripts Office au niveau de l’administrateur

Découvrez les locataires qui utilisent Office Scripts avec le journal d’audit dans le centre de conformité. Pour savoir comment utiliser cet outil, consultez [le journal d’audit dans le Centre de sécurité & conformité](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Pour rechercher qui utilise Office Scripts avec l’outil de recherche, ajoutez `.osts` le **fichier, le dossier ou le champ de site**. Cette opération recherche tous les fichiers avec l’extension de fichier Office Scripts. Si une personne de votre organisation a utilisé la fonctionnalité Office Scripts, l’activité de l’utilisateur s’affiche dans les résultats de la recherche dans le journal d’audit.

## <a name="see-also"></a>Voir aussi

- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Annuler les effets des scripts Office](../testing/undo.md)
