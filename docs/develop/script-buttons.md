---
title: Exécuter Office scripts dans Excel avec des boutons
description: Ajoutez des boutons aux classeurs qui contrôlent Office scripts dans Excel.
ms.topic: overview
ms.date: 05/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: fde34d62f9abe897a8b93195ab37a75cfc73f619
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393683"
---
# <a name="run-office-scripts-in-excel-with-buttons"></a>Exécuter Office scripts dans Excel avec des boutons

Aidez vos collègues à trouver et exécuter vos scripts en ajoutant des boutons de script à un workbook.

:::image type="content" source="../images/run-from-button.png" alt-text="Un bouton de la feuille de calcul qui exécute un script lorsque l’utilisateur clique dessus.":::

## <a name="create-script-buttons"></a>Créer des boutons de script

Avec n’importe quel script, accédez au menu **Plus d’options (...)** dans la page de détails du script ou dans le volet Office de l’Éditeur de code, puis sélectionnez **Bouton Ajouter**. Cela crée un bouton dans le workbook qui exécute le script associé lorsqu’il est sélectionné. Il partage également le script avec le workbook, de sorte que tous les personnes particulièrement autorisées à écrire sur le workbook peuvent utiliser votre automatisation utile.

La capture d’écran suivante montre la page des détails du script pour un script intitulé **Créer un tableau croisé dynamique** et l’option **Ajouter un bouton** dans le menu **Plus d’options (...)** est mise en surbrillance.

:::image type="content" source="../images/add-button.png" alt-text="Option « Bouton Ajouter » dans le menu de la page détails du script.":::

## <a name="remove-script-buttons"></a>Supprimer les boutons de script

Pour arrêter le partage d’un script via un bouton, accédez au menu **Plus d’options (...)** dans la page détails du script, puis **sélectionnez Arrêter le partage**. Cela supprime tous les boutons qui exécutent le script. La suppression d’un seul bouton supprime le script de ce bouton, même si l’opération est annulée ou si le bouton est coupé et enfoncé.

## <a name="script-buttons-with-excel-on-windows"></a>Boutons de script avec Excel sur Windows

Ces boutons de script fonctionnent également sur Windows. Créez le bouton dans Excel sur le Web et les utilisateurs sur Windows peuvent exécuter votre script en cliquant sur un bouton. Notez que vous ne pouvez pas modifier les scripts dans Excel sur Windows. Vous pouvez uniquement modifier des scripts dans Excel sur le Web.

Certaines API de scripts Office peuvent ne pas être prises en charge par Excel sur Windows, en particulier les builds plus anciennes. Il s’agit notamment d’API et d’API plus récentes pour les fonctionnalités web uniquement. Si un script contient des API non prises en charge, le script ne s’exécute pas et, à la place, le volet Des tâches **d’état** d’exécution de script affiche un message d’avertissement indiquant : « Ce script doit actuellement être exécuté sur Excel sur le Web. Ouvrez le classeur dans le navigateur, puis réessayez, ou contactez le propriétaire du script pour obtenir de l’aide. »  

> [!IMPORTANT]
> Les boutons de script nécessitent [que WebView2](/deployoffice/webview2-install) fonctionne avec Excel sur Windows. Il est installé par défaut avec les dernières versions de Excel sur le Bureau, mais si vous ne parvenez pas à cliquer sur les boutons de scripts, [visitez Télécharger le runtime WebView2](https://developer.microsoft.com/en-us/microsoft-edge/webview2/#download-section) et téléchargez le moteur de navigateur.
