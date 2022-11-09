---
title: Environnement de l’éditeur de code de scripts Office
description: Les conditions préalables et les informations d’environnement pour les scripts Office dans Excel sur le Web.
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: a5a7601285553b1da4001a1870b6120f21bf5f2c
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891252"
---
# <a name="office-scripts-code-editor-environment"></a>Environnement de l’éditeur de code de scripts Office

Les scripts Office sont écrits en TypeScript ou JavaScript et utilisent les API JavaScript de scripts Office pour interagir avec un classeur Excel. L’éditeur de code est basé sur Visual Studio Code. Par conséquent, si vous avez déjà utilisé cet environnement, vous vous sentirez comme chez vous.

> [!TIP]
> Si vous êtes familiarisé avec Visual Studio Code, vous pouvez désormais l’utiliser pour écrire des scripts. Visitez [Visual Studio Code pour les scripts Office (préversion)](../develop/vscode-for-scripts.md) pour essayer cette fonctionnalité.

## <a name="scripting-language-typescript-or-javascript"></a>Langage de script : TypeScript ou JavaScript

Les scripts Office sont écrits dans [TypeScript](https://www.typescriptlang.org/docs/home.html), qui est un ensemble de scripts [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). L’enregistreur d’actions génère du code dans TypeScript et la documentation sur les scripts Office utilise TypeScript. Étant donné que TypeScript est un sur-ensemble de JavaScript, tout code de script que vous écrivez en JavaScript fonctionne parfaitement.

Les scripts Office sont en grande partie des éléments de code autonomes. Seule une petite partie des fonctionnalités de TypeScript est utilisée. Par conséquent, vous pouvez modifier des scripts sans avoir à apprendre les subtilités de TypeScript. L’éditeur de code gère également l’installation, la compilation et l’exécution du code. Vous n’avez donc pas à vous soucier du script lui-même. Il est possible d’apprendre le langage et de créer des scripts sans connaissances préalables en programmation. Toutefois, si vous débutez dans la programmation, nous vous recommandons d’apprendre quelques notions de base avant de passer aux scripts Office :

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Les scripts Office utilisent une version spécialisée des API JavaScript Office pour les [compléments Office](/office/dev/add-ins/overview/index). Bien qu’il existe des similitudes dans les deux API, vous ne devez pas supposer que le code peut être porté entre les deux plateformes. Les différences entre les deux plateformes sont décrites dans l’article [Différences entre les scripts Office et les compléments Office](../resources/add-ins-differences.md#apis) . Vous pouvez afficher toutes les API disponibles pour votre script dans la [documentation de référence de l’API Scripts Office](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Prise en charge des bibliothèques externes

Les scripts Office ne prennent pas en charge l’utilisation de bibliothèques JavaScript tierces externes. Actuellement, vous ne pouvez pas appeler une bibliothèque autre que les API Scripts Office à partir d’un script. Vous avez toujours accès à n’importe quel [objet JavaScript intégré](../develop/javascript-objects.md), tel que [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="intellisense"></a>Intellisense

IntelliSense est un ensemble de fonctionnalités de l’éditeur de code qui vous aident à écrire du code. Il fournit la saisie semi-automatique, la mise en surbrillance des erreurs de syntaxe et la documentation de l’API inline.

IntelliSense fournit des suggestions au fur et à mesure que vous tapez, comme le texte suggéré dans Excel. Le fait d’appuyer sur la touche Tab ou Entrée insère le membre suggéré. Déclenchez IntelliSense à l’emplacement actuel du curseur en appuyant sur les touches Ctrl+Espace. Ces suggestions sont particulièrement utiles lors de l’exécution d’une méthode. La signature de méthode affichée par IntelliSense contient une liste d’arguments dont elle a besoin, le type de chaque argument, si un argument donné est obligatoire ou facultatif, et le type de retour de la méthode.

Placez le curseur sur une méthode, une classe ou un autre objet de code pour afficher plus d’informations. Pointez sur une erreur de syntaxe ou une suggestion de code, représentée par une ligne ondulée rouge ou jaune, pour afficher des suggestions sur la façon de résoudre le problème. Souvent, IntelliSense fournit une option « Correctif rapide » pour modifier automatiquement le code.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="Message d’erreur dans le texte de pointage de l’éditeur de code avec un bouton « Correctif rapide ».":::

L’Éditeur de code de scripts Office utilise le même moteur IntelliSense que Visual Studio Code. Pour en savoir plus sur la fonctionnalité, consultez [fonctionnalités IntelliSense de Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="keyboard-shortcuts"></a>Raccourcis clavier

La plupart des raccourcis clavier pour Visual Studio Code fonctionnent également dans l’Éditeur de code de scripts Office. Utilisez les fichiers PDF suivants pour en savoir plus sur les options disponibles et tirer le meilleur parti de l’éditeur de code :

- [Raccourcis clavier pour macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Raccourcis clavier pour Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Voir aussi

- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](../develop/javascript-objects.md)
- [Visual Studio Code pour les scripts Office (préversion)](../develop/vscode-for-scripts.md)
