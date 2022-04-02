---
title: environnement Office Scripts Code Editor
description: Les conditions préalables et les informations d’environnement pour Office scripts dans Excel sur le Web.
ms.date: 05/27/2021
ms.localizationpriority: medium
ms.openlocfilehash: 165365d82aa838f6651461f6edee2389c44e90b1
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585932"
---
# <a name="office-scripts-code-editor-environment"></a>environnement Office Scripts Code Editor

Office scripts sont écrits en TypeScript ou JavaScript et utilisent les API JavaScript Office Scripts pour interagir avec un Excel de travail. L’éditeur de code est basé sur Visual Studio Code, donc si vous avez déjà utilisé cet environnement auparavant, vous vous sentirez comme chez vous.

## <a name="scripting-language-typescript-or-javascript"></a>Langage de script : TypeScript ou JavaScript

Les scripts Office sont écrits dans [TypeScript](https://www.typescriptlang.org/docs/home.html), qui est un ensemble de scripts [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). L’enregistreur d’actions génère du code dans TypeScript et la documentation Office Scripts utilise TypeScript. Étant donné que TypeScript est un sur-ensemble de Code JavaScript, tout code de script que vous écrivez en JavaScript fonctionne parfaitement.

Office scripts sont en grande partie des éléments de code autonomes. Seule une petite partie des fonctionnalités de TypeScript est utilisée. Par conséquent, vous pouvez modifier des scripts sans avoir à découvrir les complexités de TypeScript. L’éditeur de code gère également l’installation, la compilation et l’exécution du code. Vous n’avez donc pas à vous soucier du script proprement dit. Il est possible d’apprendre le langage et de créer des scripts à l’insu des connaissances de programmation précédentes. Toutefois, si vous débutez dans la programmation, nous vous recommandons d’apprendre quelques principes de base avant de Office scripts :

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>API JavaScript Office Scripts

Office scripts utilisent une version spécialisée des API JavaScript Office pour Office [de recherche](/office/dev/add-ins/overview/index). Bien qu’il existe des similitudes dans les deux API, vous ne devez pas supposer que le code peut être porté entre les deux plateformes. Les différences entre les deux plateformes sont décrites dans l’article [Differences between Office Scripts and Office Add-ins](../resources/add-ins-differences.md#apis). Vous pouvez afficher toutes les API disponibles pour votre script dans la documentation de référence de [l’API Office Scripts](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Prise en charge des bibliothèques externes

Office scripts ne prend pas en charge l’utilisation de bibliothèques JavaScript tierces externes. Actuellement, vous ne pouvez pas appeler une bibliothèque autre que Office API Scripts à partir d’un script. Vous avez toujours accès à n’importe quel objet [JavaScript intégré](../develop/javascript-objects.md), tel que [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>IntelliSense

IntelliSense est un ensemble de fonctionnalités d’éditeur de code qui vous aident à écrire du code. Il fournit la mise en surbrillance automatique, la mise en surbrillance des erreurs de syntaxe et la documentation de l’API en ligne.

IntelliSense suggestions à mesure que vous tapez, similaire au texte suggéré dans Excel. Appuyer sur la touche de tabulation ou d’entrée insère le membre suggéré. Déclenchez IntelliSense à l’emplacement actuel du curseur en appuyant sur les touches Ctrl+Espace. Ces suggestions sont particulièrement utiles lors de l’exécution d’une méthode. La signature de méthode affichée par IntelliSense contient une liste d’arguments dont elle a besoin, le type de chaque argument, qu’un argument donné soit obligatoire ou facultatif, et le type de retour de la méthode.

Placez le curseur sur une méthode, une classe ou un autre objet code pour voir plus d’informations. Pointez sur une erreur de syntaxe ou une suggestion de code, représentée par une ligne rouge ou jaune, pour voir des suggestions sur la façon de résoudre le problème. Souvent, IntelliSense fournit une option « Correctif rapide » pour modifier automatiquement le code.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Message d’erreur dans le texte de pointeur de l’éditeur de code avec un bouton « Correctif rapide ».":::

L Office’éditeur de code scripts utilise le même moteur IntelliSense scripts que Visual Studio Code. Pour en savoir plus sur la fonctionnalité, consultez [Visual Studio Code fonctionnalités IntelliSense de l’équipe](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Raccourcis clavier

La plupart des raccourcis clavier pour Visual Studio Code fonctionnent également dans l’éditeur de Office scripts. Utilisez les PDF suivants pour en savoir plus sur les options disponibles et tirer le meilleur profit de l’éditeur de code :

- [Raccourcis clavier pour macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Raccourcis clavier pour Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Voir aussi

- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](../develop/javascript-objects.md)
