---
title: Scripts de Office dépannage
description: Débogage des conseils et des techniques pour Office scripts, ainsi que des ressources d’aide.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545554"
---
# <a name="troubleshoot-office-scripts"></a>Scripts de Office dépannage

Au fur et à mesure Office scripts, vous pouvez faire des erreurs. C'est bon. Vous avez les outils pour aider à trouver les problèmes et obtenir vos scripts fonctionnent parfaitement.

## <a name="types-of-errors"></a>Types d’erreurs

Office Les erreurs de script s’insurdent dans l’une des deux catégories suivantes :

* Compiler les erreurs ou les avertissements
* Erreurs de temps d’exécution

### <a name="compile-time-errors"></a>Erreurs de compilement

Les erreurs et avertissements de compilation sont d’abord affichés dans l’éditeur de code. Ceux-ci sont montrés par les soulignements rouges ondulés dans l’éditeur. Ils sont également affichés sous **l’onglet Problèmes** au bas du volet de tâche de l’éditeur de code. Le choix de l’erreur donnera plus de détails sur le problème et proposera des solutions. Les erreurs de temps de compilation doivent être traitées avant d’exécuter le script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Erreur de compilateur affichée dans le texte stationnaire de l’éditeur de code":::

Vous pouvez également voir des soulignements d’avertissement orange et des messages d’information gris. Ceux-ci indiquent des suggestions de performances ou d’autres possibilités où le script peut avoir des effets involontaires. Ces avertissements devraient être examinés de près avant de les rejeter.

### <a name="runtime-errors"></a>Erreurs de temps d’exécution

Les erreurs d’exécution se produisent en raison de problèmes logiques dans le script. Cela peut être parce qu’un objet utilisé dans le script n’est pas dans le cahier de travail, une table est formatée différemment que prévu, ou un autre léger écart entre les exigences du script et le manuel actuel. Le script suivant génère une erreur lorsqu’une feuille de travail nommée « Feuille de test » n’est pas présente.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a>Messages console

Les erreurs de compilement et d’exécution affichent les messages d’erreur dans la console lorsqu’un script s’exécute. Ils donnent un numéro de ligne où le problème a été rencontré. Gardez à l’esprit que la cause profonde de tout problème peut être une ligne de code différente de ce qui est indiqué dans la console.

L’image suivante affiche la sortie de la console pour [l’erreur `any` compilateur](../develop/typescript-restrictions.md) explicite. Notez le texte `[5, 16]` au début de la chaîne d’erreur. Cela indique que l’erreur est sur la ligne 5, à partir du caractère 16.
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La console Code Editor affichant un message d’erreur explicite « n’importe quel »":::

L’image suivante affiche la sortie de la console pour une erreur de temps d’exécution. Ici, le script tente d’ajouter une feuille de travail avec le nom d’une feuille de travail existante. Encore une fois, notez la « ligne 2 » précédant l’erreur pour montrer quelle ligne enquêter.
:::image type="content" source="../images/runtime-error-console.png" alt-text="La console Code Editor affichant une erreur de l’appel 'addWorksheet'":::

## <a name="console-logs"></a>Journaux de console

Imprimez des messages à l’écran avec `console.log` l’instruction. Ces journaux peuvent vous montrer la valeur actuelle des variables ou les chemins de code qui sont déclenchés. Pour ce faire, appelez avec `console.log` n’importe quel objet comme paramètre. Habituellement, un `string` est le type le plus facile à lire dans la console.

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

Les chaînes `console.log` transmises sont affichées dans la console de journalisation de l’éditeur de code, au bas du volet de tâche. Les journaux se trouvent sur **l’onglet Sortie,** bien que l’onglet gagne automatiquement la mise au point lorsqu’un journal est écrit.

Les journaux n’affectent pas le cahier de travail.

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a>Automatisez l’onglet n’apparaissant pas ou n’Office scripts non disponibles

Les étapes suivantes devraient aider à résoudre tous les problèmes liés à **l’onglet Automate** n’apparaissant pas dans Excel sur le Web.

1. [Assurez-vous que Microsoft 365 licence inclut Office scripts](../overview/excel.md#requirements).
1. [Vérifiez que votre navigateur est pris en charge](platform-limits.md#browser-support).
1. [Assurez-vous que les cookies tiers sont activés](platform-limits.md#third-party-cookies).
1. [Assurez-vous que votre administrateur n’a pas désactivé Office scripts dans le Microsoft 365 d’administration](/microsoft-365/admin/manage/manage-office-scripts-settings).

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a>Scripts de dépannage dans Power Automate

Pour plus d’informations spécifiques à l’exécution de scripts Power Automate, [consultez Les scripts de Office dépannage en cours d’exécution Power Automate](power-automate-troubleshooting.md).

## <a name="help-resources"></a>Ressources d’aide

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) est une communauté de développeurs prêts à aider avec les problèmes de codage. Souvent, vous serez en mesure de trouver la solution à votre problème grâce à une recherche rapide Stack Overflow. Si ce n’est pas le cas, posez votre question et étiqueter avec l’étiquette « scripts de bureau ». N’oubliez pas de mentionner que vous créez un script *Office,* pas un Office *Add-in*.

Si vous rencontrez un problème avec l’API JavaScript Office, créez un problème dans le référentiel [officedev/office-js](https://github.com/OfficeDev/office-js) GitHub.s. Les membres de l’équipe produit répondront aux problèmes et fourniront une aide supplémentaire. La création d’un problème dans le référentiel **OfficeDev/office-js** indique que vous avez trouvé une faille dans la bibliothèque d’API JavaScript Office que l’équipe produit doit traiter.

S’il y a un problème avec l’enregistreur d’action ou l’éditeur, envoyez des **commentaires via le bouton > d’aide** et de rétroaction Excel.

## <a name="see-also"></a>Voir aussi

- [Meilleures pratiques dans Office scripts](../develop/best-practices.md)
- [Limites de plate-forme avec Office scripts](platform-limits.md)
- [Améliorez les performances de vos scripts Office’argent](../develop/web-client-performance.md)
- [Scripts de Office en cours d’exécution dans PowerAutomate](power-automate-troubleshooting.md)
- [Annuler les effets des scripts Office texte](undo.md)
