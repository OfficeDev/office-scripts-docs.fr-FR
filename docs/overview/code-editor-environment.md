---
title: Office Environnement scripts Code Editor
description: Les conditions préalables et l’information sur l’environnement Office scripts en Excel sur le Web.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: aa54939826f8dda2a068df0f3fabf0fd3a2c842b
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545821"
---
# <a name="office-scripts-code-editor-environment"></a>Office Environnement scripts Code Editor

Office Les scripts sont écrits dans TypeScript ou JavaScript et utilisent les API JavaScript Office Scripts pour interagir avec un Excel de travail. L’éditeur de code est basé Visual Studio Code, donc si vous avez utilisé cet environnement avant, vous vous sentirez comme chez vous.

## <a name="scripting-language-typescript-or-javascript"></a>Langage de script : TypeScript ou JavaScript

Office Les scripts sont écrits dans [TypeScript](https://www.typescriptlang.org/docs/home.html), qui est un superset de [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). L’enregistreur d’action génère du code dans TypeScript et Office documentation scripts utilise TypeScript. Depuis TypeScript est un superset de JavaScript, tout code de script que vous écrivez dans JavaScript fonctionnera très bien.

Office Les scripts sont en grande partie des morceaux de code autonomes. Seule une petite partie des fonctionnalités de TypeScript est utilisée. Par conséquent, vous pouvez modifier des scripts sans avoir à apprendre les subtilités de TypeScript. L’éditeur de code gère également l’installation, la compilation et l’exécution du code, de sorte que vous n’avez pas besoin de vous soucier de quoi que ce soit, mais le script lui-même. Il est possible d’apprendre la langue et de créer des scripts sans connaissances préalables en programmation. Toutefois, si vous êtes nouveau dans la programmation, nous vous recommandons d’apprendre certains principes fondamentaux avant de procéder Office scripts :

[!INCLUDE [Preview note](../includes/coding-basics-references.md)]

## <a name="office-scripts-javascript-api"></a>Office Scripts JavaScript API

Office Les scripts utilisent une version spécialisée des Office API JavaScript pour [Office Add-ins](/office/dev/add-ins/overview/index). Bien qu’il existe des similitudes dans les deux API, vous ne devez pas supposer que le code peut être porté entre les deux plates-formes. Les différences entre les deux plateformes sont décrites dans [les différences entre Office scripts et Office’article Add-ins.](../resources/add-ins-differences.md#apis) Vous pouvez afficher toutes les API disponibles pour votre script dans la [documentation de référence Office Scripts API](/javascript/api/office-scripts/overview).

## <a name="external-library-support"></a>Support externe de la bibliothèque

Office Les scripts ne prend pas en charge l’utilisation de bibliothèques JavaScript externes et tierces. Actuellement, vous ne pouvez pas appeler une bibliothèque autre que les Office scripts d’un script. Vous avez toujours accès à [n’importe quel objet JavaScript intégré,](../develop/javascript-objects.md)comme [les mathématiques.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math)

## <a name="intellisense"></a>IntelliSense

IntelliSense est une fonctionnalité d’éditeur de code qui aide à prévenir les fautes de frappe et les erreurs de syntaxe lorsque vous modifiez votre script. Il affiche les noms d’objets et de champs possibles au fur et à mesure que vous tapez, ainsi que la documentation en ligne pour chaque API.

L Excel éditeur de code utilise le même IntelliSense que Visual Studio Code. Pour en savoir plus sur la fonctionnalité, [visitez Visual Studio Code’IntelliSense de l’équipe.](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features)

## <a name="keyboard-shortcuts"></a>Raccourcis clavier

La plupart des raccourcis clavier pour Visual Studio Code également dans le Office Scripts Code Editor. Utilisez les FICHIERS PDF suivants pour en savoir plus sur les options disponibles et tirer le meilleur parti de l’éditeur de code :

- [Raccourcis clavier pour macOS](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-macos.pdf).
- [Raccourcis clavier pour Windows](https://code.visualstudio.com/shortcuts/keyboard-shortcuts-windows.pdf).

## <a name="see-also"></a>Voir aussi

- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](../develop/javascript-objects.md)
