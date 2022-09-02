---
title: Stockage et propriété de fichiers Office Scripts
description: Informations sur la façon dont les scripts Office sont stockés dans Microsoft OneDrive et transférés entre propriétaires.
ms.date: 08/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 573f65f299c29b4f481c9a2e23ebe7e36181706b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572506"
---
# <a name="office-scripts-file-storage-and-ownership"></a>Stockage et propriété de fichiers Office Scripts

Les scripts Office sont stockés sous forme de fichiers **.osts** dans votre dossier Microsoft OneDrive ou SharePoint. Ils sont stockés séparément d’un classeur. Pour permettre aux utilisateurs qui se trouvent en dehors du site SharePoint d’accéder au script, [partagez le script avec un classeur Excel](excel.md#share-office-scripts). Cela signifie que vous liez le script avec le fichier, et non l’attachez. Quiconque a accès au fichier Excel sera également en mesure d’afficher, d’exécuter ou de faire une copie du script.

Excel reconnaît et exécute un script uniquement s’il se trouve dans votre dossier OneDrive, un dossier Sharepoint ou partagé avec le classeur.

## <a name="onedrive"></a>OneDrive

Le comportement par défaut est que les scripts Office sont stockés dans votre OneDrive. Les fichiers **.osts** se trouvent dans le dossier **/Documents/Scripts Office/** . Toutes les modifications apportées à ces fichiers **.osts** , telles que le changement de nom ou la suppression de fichiers, sont répercutées dans l’Éditeur de code et la galerie de scripts.

Les scripts partagés avec l’un de vos classeurs restent dans le OneDrive du créateur du script. Ils ne sont copiés dans aucun de vos dossiers locaux ou OneDrive lorsque vous exécutez le script partagé dans Excel. Le bouton **Créer une copie** de l’éditeur de code enregistre une copie distincte du script dans votre OneDrive. Les modifications apportées à la copie n’affectent pas le script d’origine.

Sauf si vous partagez vos scripts personnels, personne d’autre ne peut y accéder. Vos paramètres OneDrive contrôlent l’accès partagé et les autorisations pour tous les fichiers **.osts** de script, indépendamment des paramètres Excel. Les scripts ne peuvent pas être liés à partir d’un disque local ou d’emplacements cloud personnalisés.

## <a name="sharepoint"></a>SharePoint

Les scripts Office enregistrés sur un site SharePoint appartiennent à votre équipe. Vous et les membres de votre organisation disposant de l’accès approprié peuvent exécuter et modifier des scripts à partir de SharePoint. Ces scripts s’affichent également dans la galerie de scripts de l’onglet **Automatiser** .

Pour charger un script à partir de SharePoint, accédez à **Tous les scripts** et sélectionnez **Afficher d’autres scripts** en bas de la liste. Cela fait apparaître un sélecteur de fichiers dans lequel vous pouvez choisir **des fichiers .osts** à partir de n’importe quel site SharePoint auquel vous avez accès. Notez que les scripts de SharePoint que vous avez déjà ouverts s’affichent dans la liste des scripts récents.

Pour enregistrer un script dans SharePoint, accédez au menu **Plus d’options (...)** et **sélectionnez Enregistrer sous**. Cela ouvre un sélecteur de fichiers dans lequel vous pouvez sélectionner des dossiers dans votre site SharePoint. L’enregistrement dans un nouvel emplacement crée une copie du script à cet emplacement. La version d’origine se trouve toujours sur votre oneDrive ou un autre emplacement SharePoint.

> [!IMPORTANT]
> Les scripts avec [des appels externes](../develop/external-calls.md) ne peuvent pas être exécutés à partir de SharePoint. Vous recevrez une erreur indiquant que « Les appels d’accès réseau ne sont pas pris en charge pour l’instant pour les scripts enregistrés sur un site SharePoint ».

> [!IMPORTANT]
> Power Automate ne prend **pas** en charge les scripts stockés sur SharePoint pour l’instant.

## <a name="restore-deleted-scripts"></a>Restaurer des scripts supprimés

Lorsque vous supprimez un script dans Excel, il est envoyé à votre corbeille OneDrive ou SharePoint. Pour restaurer un script supprimé, suivez les étapes répertoriées dans [Comment récupérer des éléments manquants, supprimés ou endommagés dans SharePoint et OneDrive pour le travail ou l’école](https://support.microsoft.com/office/how-to-recover-missing-deleted-or-corrupted-items-in-sharepoint-and-onedrive-for-work-or-school-3d748edf-c072-46c9-81a4-4989056ebc87). La restauration d’un fichier **.osts** le renvoie à la liste **Tous les scripts** .

Un script supprimé n’est pas partagé avec le classeur. Lorsque vous restaurez un script, il ne conserve **pas** son accès au script. Vous devrez partager à nouveau le script.

Les scripts restaurés fonctionnent toujours comme prévu avec les flux Power Automate. Vous n’avez pas besoin de recréer le connecteur de flux.

## <a name="file-ownership-and-retention"></a>Propriété et rétention des fichiers

Les scripts Office suivent les stratégies de rétention et de suppression spécifiées par Microsoft OneDrive et Microsoft SharePoint. Pour savoir comment gérer les scripts qui ont été créés et partagés par un utilisateur supprimé de votre organisation, consultez [En savoir plus sur la rétention pour SharePoint et OneDrive](/microsoft-365/compliance/retention-policies-sharepoint?view=o365-worldwide&preserve-view=true).

Pendant la modification, les fichiers sont temporairement stockés dans le navigateur. Vous devez enregistrer le script avant de fermer la fenêtre Excel pour l’enregistrer à l’emplacement OneDrive. N’oubliez pas d’enregistrer le fichier après les modifications, sinon ces modifications se trouveront uniquement dans la version du fichier du navigateur.

## <a name="audit-office-scripts-usage-at-the-admin-level"></a>Auditer l’utilisation des scripts Office au niveau de l’administrateur

Découvrez qui utilise les scripts Office dans votre organisation avec le journal d’audit du Centre de conformité. Pour plus d’informations sur le journal d’audit, consultez [le journal d’audit du Centre de sécurité & conformité](/microsoft-365/compliance/search-the-audit-log-in-security-and-compliance?view=o365-worldwide&preserve-view=true#search-the-audit-log).

Pour auditer spécifiquement l’activité liée aux scripts Office en tant qu’administrateur, procédez comme suit.

1. Dans une fenêtre de navigateur InPrivate (ou Incognito ou un autre mode de suivi limité spécifique au navigateur), ouvrez et connectez-vous au [Centre de conformité](https://compliance.microsoft.com/).
1. Accédez à la page **Audit** .
1. *(Une seule fois)* Sous l’onglet **Rechercher** , **sélectionnez Démarrer l’enregistrement de l’activité de l’utilisateur et de l’administrateur**.

    > [!IMPORTANT]
    > Il peut s’écouler une heure ou deux après l’activation de l’enregistrement avant que toutes les activités sur le locataire soient enregistrées.

1. Définissez les options de recherche souhaitées et **appuyez sur Recherche**. **Filtrer les activités** pour **exécuter le script sur le classeur** pour voir chaque fois qu’un script a été exécuté. Vous pouvez également filtrer le **champ Fichier, dossier ou site** sur `.osts`. Cela révèle qui, dans votre organisation, crée ou modifie des scripts.

    :::image type="content" source="../images/audit-log-example.png" alt-text="Quelques lignes de résultats de recherche dans le journal d’audit, notamment l’action « Exécuter le script sur le classeur » et le chargement et la modification d’un fichier .osts.":::

## <a name="see-also"></a>Voir aussi

- [Partager des scénarios de bureau en Excel pour le Web](https://support.microsoft.com/office/226eddbc-3a44-4540-acfe-fccda3d1122b)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Paramètres de Office Scripts dans M365](/microsoft-365/admin/manage/manage-office-scripts-settings)
- [Annuler les effets des scripts Office](../testing/undo.md)
