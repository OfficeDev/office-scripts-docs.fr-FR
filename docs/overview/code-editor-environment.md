---
title: Environnement de l’éditeur de code des scripts Office
description: Les conditions préalables et les informations d’environnement pour les scripts Office dans Excel sur le Web.
ms.date: 04/24/2020
localization_priority: Normal
ms.openlocfilehash: efe6ddbe39a1ea3850b4dc6fea0fa885b80c0c28
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878667"
---
# <a name="office-scripts-code-editor-environment"></a>Environnement de l’éditeur de code des scripts Office

Les scripts Office sont écrits en écriture manuscrite [ou en JavaScript](#scripting-language-typescript-or-javascript) et utilisent les [API JavaScript de scripts Office](#office-scripts-javascript-api) pour interagir avec un classeur Excel.

## <a name="scripting-language-typescript-or-javascript"></a>Langage de script : écriture ou JavaScript

Les scripts Office sont écrits [en écriture](https://www.typescriptlang.org/docs/home.html) manuscrite ou en [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). L’enregistreur d’actions génère du code dans la machine à écrire (qui est un sur-ensemble de JavaScript). La documentation sur les scripts Office utilise la machine à écrire, mais si vous êtes plus à l’aise avec JavaScript, vous pouvez l’utiliser à la place.

Les scripts Office sont principalement des portions de code autonomes. Seule une petite partie de la fonctionnalité d’écriture est utilisée. Par conséquent, vous pouvez modifier les scripts sans avoir à vous familiariser avec les subtilités de la machine à écrire. L’éditeur de code gère également l’installation, la compilation et l’exécution du code, de sorte que vous n’avez pas à vous soucier de tout, à l’exception du script lui-même. Il est possible d’apprendre le langage et de créer des scripts sans connaissance préalable de la programmation. Toutefois, si vous débutez dans la programmation, nous vous recommandons d’apprendre des notions de base avant de procéder à des scripts Office :

- Découvrez les notions de base de JavaScript. Vous devez être familiarisé avec les concepts comme les variables, le flux de contrôle, les fonctions et les types de données. [Mozilla offre un didacticiel performant et complet sur JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- En savoir plus sur les types dans la machine à écrire. La méthode dactylographié s’appuie sur JavaScript en s’assurant au moment de la compilation que les types corrects sont utilisés pour les appels et les affectations de méthodes. La documentation de la machine à écrire sur les [interfaces](https://www.typescriptlang.org/docs/handbook/interfaces.html), les [classes](https://www.typescriptlang.org/docs/handbook/classes.html), l' [inférence de type](https://www.typescriptlang.org/docs/handbook/type-inference.html)et la [compatibilité des types](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) seront les plus utiles.

## <a name="office-scripts-javascript-api"></a>API JavaScript pour les scripts Office

Les scripts Office utilisent une version spécialisée des API JavaScript pour Office pour les [Compléments Office](/office/dev/add-ins/overview/index). Bien qu’il existe des similitudes entre les deux API, vous ne devez pas supposer que le code peut être transféré entre les deux plateformes. Les différences entre les deux plateformes sont décrites dans l’article [différences entre les scripts Office et les compléments Office](../resources/add-ins-differences.md#apis) . Vous pouvez afficher toutes les API disponibles pour votre script dans la [documentation de référence de l’API des scripts Office](/javascript/api/office-scripts/overview).

## <a name="intellisense"></a>Remplissage

IntelliSense est une fonctionnalité d’éditeur de code qui permet d’éviter les fautes de frappe et de syntaxe lors de la modification de votre script. Il affiche les noms d’objet et de champ possibles en fonction de votre type, ainsi que la documentation en ligne pour chaque API.

L’éditeur de code Excel utilise le même moteur IntelliSense que Visual Studio code. Pour en savoir plus sur la fonctionnalité, consultez la rubrique [fonctionnalités IntelliSense de Visual Studio code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="external-library-support"></a>Prise en charge des bibliothèques externes

Les scripts Office ne prennent pas en charge l’utilisation de bibliothèques JavaScript tierces externes. Vous ne pouvez actuellement pas appeler une bibliothèque autre que les API de scripts Office à partir d’un script. Vous avez toujours accès à un [objet JavaScript intégré](../develop/javascript-objects.md), tel que [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="see-also"></a>Voir aussi

- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](../develop/javascript-objects.md)
