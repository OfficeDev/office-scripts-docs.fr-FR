---
title: Résoudre les problèmes liés aux scripts Office
description: Conseils et techniques de débogage pour les scripts Office, ainsi que des ressources d’aide.
ms.date: 10/05/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4fe4a9b17d51d078403d1a46abed774d38eeaa80
ms.sourcegitcommit: 64d506257bee282fb01aedbf4d090781b06e4900
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/07/2022
ms.locfileid: "68495466"
---
# <a name="troubleshoot-office-scripts"></a>Résoudre les problèmes liés aux scripts Office

Lorsque vous développez des scripts Office, vous pouvez faire des erreurs. C'est bon. Vous disposez des outils nécessaires pour vous aider à trouver les problèmes et à faire fonctionner parfaitement vos scripts.

> [!NOTE]
> Pour obtenir des conseils de dépannage spécifiques aux scripts Office avec Power Automate, consultez [Résolution des problèmes liés aux scripts Office en cours d’exécution dans Power Automate](power-automate-troubleshooting.md).

## <a name="types-of-errors"></a>Types d’erreurs

Les erreurs de scripts Office se répartissent dans l’une des deux catégories suivantes :

* Erreurs ou avertissements au moment de la compilation
* Erreurs d’exécution

### <a name="compile-time-errors"></a>Erreurs de compilation

Les erreurs et avertissements au moment de la compilation sont initialement affichés dans l’Éditeur de code. Ceux-ci sont affichés par les soulignements ondulés rouges dans l’éditeur. Ils sont également affichés sous l’onglet **Problèmes** en bas du volet Office de l’Éditeur de code. La sélection de l’erreur donne plus de détails sur le problème et suggère des solutions. Les erreurs de compilation doivent être résolues avant d’exécuter le script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Erreur du compilateur affichée dans le texte de pointage de l’éditeur de code.":::

Vous pouvez également voir des soulignements d’avertissement orange et des messages d’information gris. Elles indiquent des suggestions de performances ou d’autres possibilités où le script peut avoir des effets involontaires. Ces avertissements doivent être examinés de près avant de les ignorer.

### <a name="runtime-errors"></a>Erreurs d’exécution

Des erreurs d’exécution se produisent en raison de problèmes logiques dans le script. Cela peut être dû au fait qu’un objet utilisé dans le script n’est pas dans le classeur, qu’une table est mise en forme différemment de ce qui était prévu ou qu’il existe une légère différence entre les exigences du script et le classeur actuel. Le script suivant génère une erreur lorsqu’une feuille de calcul nommée « TestSheet » n’est pas présente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Messages de la console

Les erreurs de compilation et d’exécution affichent des messages d’erreur dans la console lors de l’exécution d’un script. Ils donnent un numéro de ligne où le problème a été rencontré. N’oubliez pas que la cause racine d’un problème peut être une ligne de code différente de celle indiquée dans la console.

L’image suivante montre la sortie de la console pour l’erreur [explicite `any`](../develop/typescript-restrictions.md) du compilateur. Notez le texte `[5, 16]` au début de la chaîne d’erreur. Cela indique que l’erreur se trouve sur la ligne 5, en commençant par le caractère 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Console de l’éditeur de code affichant un message d’erreur « any » explicite.":::

L’image suivante montre la sortie de la console pour une erreur d’exécution. Ici, le script tente d’ajouter une feuille de calcul portant le nom d’une feuille de calcul existante. Là encore, notez la « ligne 2 » qui précède l’erreur pour afficher la ligne à examiner.
:::image type="content" source="../images/runtime-error-console.png" alt-text="Console de l’éditeur de code affichant une erreur à partir de l’appel « addWorksheet ».":::

## <a name="console-logs"></a>Journaux de la console

Imprimez les messages à l’écran avec l’instruction `console.log` . Ces journaux peuvent vous montrer la valeur actuelle des variables ou les chemins de code qui sont déclenchés. Pour ce faire, appelez `console.log` n’importe quel objet en tant que paramètre. En règle générale, un `string` type est le plus simple à lire dans la console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Les chaînes transmises sont `console.log` affichées dans la console de journalisation de l’Éditeur de code, en bas du volet Office. Les journaux se trouvent sous l’onglet **Sortie** , bien que l’onglet gagne automatiquement le focus lorsqu’un journal est écrit.

Les journaux n’affectent pas le classeur.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>L’onglet Automatiser n’apparaît pas ou les scripts Office ne sont pas disponibles

Les étapes suivantes doivent vous aider à résoudre les problèmes liés à l’onglet **Automatiser** qui n’apparaît pas dans Excel sur le Web.

1. [Assurez-vous que votre licence Microsoft 365 inclut des scripts Office](../overview/excel.md#requirements).
1. [Vérifiez que votre navigateur est pris en charge](platform-limits.md#browser-support).
1. [Vérifiez que les cookies tiers sont activés](platform-limits.md#third-party-cookies).
1. [Vérifiez que votre administrateur n’a pas désactivé les scripts Office dans le Centre d'administration Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).
1. Vérifiez que vous n’êtes pas connecté en tant qu’utilisateur externe ou invité à votre locataire.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

> [!NOTE]
> Il existe un problème connu qui empêche les scripts stockés dans SharePoint d’apparaître toujours dans la liste récemment utilisée. Cela se produit lorsque votre administrateur désactive Exchange Web Services (EWS). Vos scripts SharePoint sont toujours accessibles et utilisables par le biais de la boîte de dialogue de fichier.

## <a name="help-resources"></a>Ressources d’aide

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) est une communauté de développeurs prêts à aider à résoudre les problèmes de codage. Souvent, vous serez en mesure de trouver la solution à votre problème par le biais d’une recherche stack overflow rapide. Si ce n’est pas le cas, posez votre question et étiquetez-la avec la balise « office-scripts ». Veillez à mentionner que vous créez un *script* Office, et non un *complément* Office.

## <a name="see-also"></a>Voir aussi

- [Meilleures pratiques en matière de scripts Office](../develop/best-practices.md)
- [Limites de plateforme avec les scripts Office](platform-limits.md)
- [Améliorer les performances de vos scripts Office](../develop/web-client-performance.md)
- [Résoudre les problèmes liés aux scripts Office en cours d’exécution dans PowerAutomate](power-automate-troubleshooting.md)
- [Annuler les effets des scripts Office](undo.md)
