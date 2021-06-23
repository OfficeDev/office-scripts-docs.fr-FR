---
title: Résoudre les problèmes Office scripts
description: Conseils et techniques de débogage pour Office scripts, ainsi que des ressources d’aide.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 251ad72588422a86c52c81666164c2c4bd79bdb5
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074647"
---
# <a name="troubleshoot-office-scripts"></a>Résoudre les problèmes Office scripts

Lorsque vous développez Office scripts, vous pouvez faire des erreurs. C'est bon. Vous avez les outils nécessaires pour trouver les problèmes et faire fonctionner parfaitement vos scripts.

## <a name="types-of-errors"></a>Types d’erreurs

Office Les erreurs de script se classent dans l’une des deux catégories suivantes :

* Erreurs ou avertissements au moment de la compilation
* Erreurs d’runtime

### <a name="compile-time-errors"></a>Erreurs au moment de la compilation

Les erreurs et avertissements au moment de la compilation sont initialement affichés dans l’éditeur de code. Ces éléments sont affichés par les soulignements ondulés rouges dans l’éditeur. Ils sont également affichés sous l’onglet **Problèmes** en bas du volet Des tâches de l’Éditeur de code. La sélection de l’erreur donne plus de détails sur le problème et suggère des solutions. Les erreurs de compilation doivent être traitées avant l’exécution du script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Erreur de compilateur affichée dans le texte de pointeur de l’éditeur de code.":::

Vous pouvez également voir des soulignements d’avertissement orange et des messages d’information gris. Celles-ci indiquent des suggestions de performances ou d’autres possibilités dans le cas où le script peut avoir des effets involontaires. Ces avertissements doivent être examinés attentivement avant de les ignorer.

### <a name="runtime-errors"></a>Erreurs d’runtime

Les erreurs d’utilisation se produisent en raison de problèmes logiques dans le script. Cela peut être dû au fait qu’un objet utilisé dans le script ne se trouve pas dans le workbook, qu’un tableau est formaté différemment des prévisions ou qu’il existe une légère différence entre les exigences du script et le workbook actuel. Le script suivant génère une erreur lorsqu’une feuille de calcul nommée « TestSheet » n’est pas présente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Messages de la console

Les erreurs de compilation et d’runtime affichent des messages d’erreur dans la console lorsqu’un script s’exécute. Ils donnent un numéro de ligne où le problème s’est produits. N’oubliez pas que la cause première d’un problème peut être une ligne de code différente de ce qui est indiqué dans la console.

L’image suivante montre la sortie de la console pour [l’erreur `any` ](../develop/typescript-restrictions.md) explicite du compilateur. Notez le texte `[5, 16]` au début de la chaîne d’erreur. Cela indique que l’erreur se trouve sur la ligne 5, en commençant au caractère 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="Console de l’éditeur de code affichant un message d’erreur explicite « tout ».":::

L’image suivante montre la sortie de la console pour une erreur d’runtime. Ici, le script tente d’ajouter une feuille de calcul avec le nom d’une feuille de calcul existante. Là encore, notez la « ligne 2 » précédant l’erreur pour afficher la ligne à examiner.
:::image type="content" source="../images/runtime-error-console.png" alt-text="Console de l’éditeur de code affichant une erreur à partir de l’appel « addWorksheet ».":::

## <a name="console-logs"></a>Journaux de la console

Imprime les messages à l’écran avec `console.log` l’instruction. Ces journaux peuvent vous montrer la valeur actuelle des variables ou les chemins de code qui sont déclenchés. Pour ce faire, `console.log` appelez avec n’importe quel objet en tant que paramètre. En règle générale, `string` il s’agit du type le plus simple à lire dans la console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Les chaînes transmises sont affichées dans la console de journalisation de l’éditeur de code, en `console.log` bas du volet Des tâches. Les journaux se  trouvent sous l’onglet Sortie, bien que l’onglet soit automatiquement mis au point lors de l’écriture d’un journal.

Les journaux n’affectent pas le workbook.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Automatiser l’onglet qui n’apparaît pas ou Office scripts indisponibles

Les étapes suivantes doivent vous aider à résoudre les problèmes liés à l’onglet **Automatiser** qui n’apparaît pas dans Excel sur le Web.

1. [Assurez-vous que votre licence Microsoft 365 inclut Office scripts.](../overview/excel.md#requirements)
1. [Vérifiez que votre navigateur est pris en charge.](platform-limits.md#browser-support)
1. [Assurez-vous que les cookies tiers sont activés.](platform-limits.md#third-party-cookies)
1. [Assurez-vous que votre administrateur n’a pas désactivé Office scripts dans le Centre d’administration Microsoft 365](/microsoft-365/admin/manage/manage-office-scripts-settings).

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Résoudre les problèmes de scripts dans Power Automate

Pour plus d’informations sur l’exécution de scripts Power Automate, voir Résolution des problèmes Office [scripts en](power-automate-troubleshooting.md)cours d’exécution dans Power Automate .

## <a name="help-resources"></a>Ressources d’aide

[Stack Overflow est](https://stackoverflow.com/questions/tagged/office-scripts) une communauté de développeurs prêts à vous aider avec les problèmes de codage. Souvent, vous serez en mesure de trouver la solution à votre problème par le biais d’une recherche rapide de stack overflow. Si ce n’est pas le cas, posez votre question et marquez-la avec la balise « office-scripts ». N’oubliez pas de mentionner que vous créez un *script* Office, et non un *Office.*

Pour envoyer une demande de fonctionnalité pour Office Scripts, publiez votre idée sur notre [page Voix](https://excel.uservoice.com/forums/274580-excel-for-the-web?category_id=143439)utilisateur ou, si la demande de fonctionnalité existe déjà, ajoutez votre vote pour cette demande. N’oubliez pas de déposer la demande sous Excel sur le Web dans la catégorie « Macros, scripts et macros » .

En cas de problème avec l’enregistreur d’actions ou l’éditeur, n’hésitez pas à nous le faire savoir. Dans le **menu** ... de l’Éditeur  de code, sélectionnez le bouton Envoyer des commentaires pour partager les problèmes.

:::image type="content" source="../images/code-editor-feedback.png" alt-text="Menu de dépassement de l’Éditeur de code avec le bouton Envoyer des commentaires.":::

## <a name="see-also"></a>Voir aussi

- [Meilleures pratiques en matière de scripts Office](../develop/best-practices.md)
- [Limites de plateforme avec Office scripts](platform-limits.md)
- [Améliorer les performances de vos scripts Office de gestion](../develop/web-client-performance.md)
- [Résoudre les Office scripts en cours d’exécution dans PowerAutomate](power-automate-troubleshooting.md)
- [Annuler les effets des scripts Office](undo.md)
