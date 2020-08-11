---
title: Dépannage de Office Scripts
description: Débogage des conseils et techniques pour les scripts Office, ainsi que des ressources d’aide.
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 00727b497d49a2d1d3f9c61e259b8d8d75028a59
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616681"
---
# <a name="troubleshooting-office-scripts"></a>Dépannage de Office Scripts

Lorsque vous développez des scripts Office, vous pouvez faire des erreurs. C'est bon. Nous disposons d’outils qui permettent de trouver les problèmes et de faire fonctionner vos scripts parfaitement.

## <a name="console-logs"></a>Journaux de console

Parfois, lors de la résolution des problèmes, vous voudrez imprimer des messages à l’écran. Ces éléments peuvent vous indiquer la valeur actuelle des variables ou les chemins d’accès de code déclenchés. Pour ce faire, consignez le texte dans la console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Les chaînes transmises `console.log` s’afficheront dans la console de journalisation de l’éditeur de code. Pour activer la console, appuyez sur le bouton de **sélection** et sélectionnez **logs...**

Les journaux n’ont pas d’incidence sur le classeur.

## <a name="error-messages"></a>Messages d’erreur

Lorsque votre script Excel rencontre un problème, il génère une erreur. Un message contextuel s’affiche pour vous demander si vous souhaitez **afficher les journaux**. Appuyez sur ce bouton pour ouvrir la console et afficher les erreurs éventuelles.

## <a name="automate-tab-not-appearing"></a>L’onglet automatiser n’apparaît pas

Les étapes suivantes doivent vous aider à résoudre les problèmes liés à l’onglet **automatiser** qui n’apparaît pas dans Excel pour le Web.

1. Assurez-vous [que votre licence Microsoft 365 inclut des scripts Office](../overview/excel.md#requirements).
1. [Demandez à votre administrateur d’activer la fonctionnalité](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf).
1. [Vérifiez que votre navigateur est pris en charge](platform-limits.md#browser-support).
1. [Vérifiez que les cookies tiers sont activés](platform-limits.md#third-party-cookies).

## <a name="help-resources"></a>Ressources d’aide

Le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts) est une communauté de développeurs souhaitant aider à coder les problèmes. Souvent, vous pouvez trouver la solution à votre problème via une recherche de débordement de pile rapide. Si ce n’est pas le cas, posez votre question et marquez-la à l’aide de la balise « Office-scripts ». N’oubliez pas de mentionner que vous créez un *script*Office, et non un *complément*Office.

Si vous rencontrez un problème avec l’API JavaScript pour Office, créez un problème dans le référentiel GitHub [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) . Les membres de l’équipe produit répondront aux problèmes et fourniront de l’aide. La création d’un problème dans le référentiel **OfficeDev/Office-js** indique que vous avez trouvé un défaut dans la bibliothèque de l’API JavaScript Office que l’équipe produit doit résoudre.

En cas de problème avec l’enregistreur d’actions ou l’éditeur, envoyez des commentaires via le bouton **d’aide > commentaires** dans Excel.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Scripts de base pour les scripts Office dans Excel sur le Web](../develop/scripting-fundamentals.md)
- [Limites des plateformes avec les scripts Office](platform-limits.md)
- [Améliorer les performances de vos scripts Office](../develop/web-client-performance.md)
- [Annuler les effets d’un script Office](undo.md)
