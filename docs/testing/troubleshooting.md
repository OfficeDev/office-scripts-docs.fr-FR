---
title: Résolution des problèmes liés aux scripts Office
description: Débogage des conseils et techniques pour les scripts Office, ainsi que des ressources d’aide.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700204"
---
# <a name="troubleshooting-office-scripts"></a>Résolution des problèmes liés aux scripts Office

Lorsque vous développez des scripts Office, vous pouvez faire des erreurs. C'est bon. Nous disposons d’outils qui permettent de trouver les problèmes et de faire fonctionner vos scripts parfaitement.

## <a name="console-logs"></a>Journaux de console

Parfois, lors de la résolution des problèmes, vous voudrez imprimer des messages à l’écran. Ces éléments peuvent vous indiquer la valeur actuelle des variables ou les chemins d’accès de code déclenchés. Pour ce faire, consignez le texte dans la console.

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> N’oubliez pas `load` d’utiliser les `sync` données de feuille de calcul et le classeur avant de consigner les propriétés de l’objet.

Les chaînes transmises`console.log` s’afficheront dans la console de journalisation de l’éditeur de code. Pour activer la console, appuyez sur le bouton de **sélection** et sélectionnez **logs...**

Les journaux n’ont pas d’incidence sur le classeur.

## <a name="error-messages"></a>Messages d’erreur

Lorsque votre script Excel rencontre un problème, il génère une erreur. Un message contextuel s’affiche pour vous demander si vous souhaitez **afficher les journaux**. Appuyez sur ce bouton pour ouvrir la console et afficher les erreurs éventuelles.

## <a name="help-resources"></a>Ressources d’aide

Le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts) est une communauté de développeurs souhaitant aider à coder les problèmes. Souvent, vous pouvez trouver la solution à votre problème via une recherche de débordement de pile rapide. Si ce n’est pas le cas, posez votre question et marquez-la à l’aide de la balise « Office-scripts ». N’oubliez pas de mentionner que vous créez un *script*Office, et non un *complément*Office.

Si vous rencontrez un problème avec l’API JavaScript pour Office, créez un problème dans le référentiel GitHub [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) . Les membres de l’équipe produit répondront aux problèmes et fourniront de l’aide. La création d’un problème dans le référentiel **OfficeDev/Office-js** indique que vous avez trouvé un défaut dans la bibliothèque de l’API JavaScript Office que l’équipe produit doit résoudre.

En cas de problème avec l’enregistreur d’actions ou l’éditeur, envoyez des commentaires via le bouton **d’aide > commentaires** dans Excel.

## <a name="see-also"></a>Voir aussi

- [Scripts Office dans Excel sur le Web](../overview/excel.md)
- [Scripts de base pour les scripts Office dans Excel sur le Web](../develop/scripting-fundamentals.md)
- [Annuler les effets d’un script Office](undo.md)
