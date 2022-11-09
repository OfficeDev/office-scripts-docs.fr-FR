---
title: Visual Studio Code pour les scripts Office (préversion)
description: Comment configurer l’éditeur de code de scripts Office pour se connecter à VS Code pour le web.
ms.date: 11/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: fd9dd417610c8ad64fbd3fc50048ce56afdb4e28
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892034"
---
# <a name="visual-studio-code-for-office-scripts-preview"></a>Visual Studio Code pour les scripts Office (préversion)

[Visual Studio Code pour le web](https://vscode.dev/) permet aux utilisateurs de modifier n’importe quoi depuis n’importe où. Connectez votre expérience de scripts Office à cet éditeur de code populaire pour commencer à créer des scripts en dehors du classeur.

:::image type="content" source="../images/vscode-script-editor.png" alt-text="Fenêtre Excel sur le Web avec l’éditeur de code ouvert en regard d’une fenêtre VS Code sur le web avec un script ouvert.":::

Visual Studio Code présente quelques avantages par rapport à l’éditeur de code intégré.

- Modification en plein écran ! Votre script n’a plus besoin de partager l’espace d’écran avec le classeur.
- Modifiez plusieurs scripts à la fois ! Basculez rapidement entre les scripts pour partager du code à partir de vos autres automatisations.
- Extensions! Utilisez vos extensions VS Code préférées pour la vérification orthographique, la mise en forme et tout ce qui vous aide à accomplir le travail.

> [!NOTE]
> Cette fonctionnalité est en préversion. Il est susceptible d’être modifié en fonction des commentaires. Si vous rencontrez des problèmes, signalez-les via le bouton **Commentaires** dans Excel. Voici les problèmes connus liés à la version actuelle de la fonctionnalité.
>
> - Visual Studio Code peut uniquement être connecté aux scripts Office via Excel sur le Web.
> - Cette connexion Office Scripts est disponible uniquement avec les clients Excel en anglais.

## <a name="connect-visual-studio-code-to-office-scripts"></a>Connecter Visual Studio Code aux scripts Office

Suivez ces étapes ponctuelles pour connecter Visual Studio Code et Excel sur le Web.

1. Ouvrez **l’Éditeur de code** de scripts Office.
2. Sous le menu **Plus d’options (...)** , sélectionnez **Paramètres de l’éditeur**.
3. Sélectionnez **(préversion) Connexion Visual Studio Code**.

:::image type="content" source="../images/vscode-enable-option.png" alt-text="Volet Office paramètres de l’éditeur affichant une case à cocher intitulée Connexion Visual Studio Code.":::

Vous pouvez maintenant modifier et exécuter vos scripts à partir de Visual Studio Code. Dans n’importe quel script, accédez au menu **Plus d’options (...)** et sélectionnez **Ouvrir dans VS Code**.

:::image type="content" source="../images/vscode-open-option.png" alt-text="L’option Ouvrir dans VS Code est sélectionnée dans une liste en regard d’un script ouvert.":::

## <a name="see-also"></a>Voir aussi

- [Environnement de l’éditeur de code de scripts Office](../overview/code-editor-environment.md)
- [Visual Studio Code pour le web (documentation)](https://code.visualstudio.com/docs/editor/vscode-web)
