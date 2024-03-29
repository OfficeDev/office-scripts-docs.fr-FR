---
title: Meilleures pratiques en matière de scripts Office
description: Comment éviter les problèmes courants et écrire des Office scripts fiables qui peuvent gérer des données ou des entrées inattendues.
ms.date: 12/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 689196e1a0ca70c999ec8048de64190cbfe75581
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585764"
---
# <a name="best-practices-in-office-scripts"></a>Meilleures pratiques en matière de scripts Office

Ces modèles et pratiques sont conçus pour aider vos scripts à s’exécuter correctement à chaque fois. Utilisez-les pour éviter les pièges courants lorsque vous commencez à automatiser Excel flux de travail.

## <a name="use-the-action-recorder-to-learn-new-features"></a>Utiliser l’enregistreur d’actions pour découvrir de nouvelles fonctionnalités

Excel fait beaucoup d’choses. La plupart d’entre eux peuvent être scriptés. L’enregistreur d’actions enregistre Excel actions et les traduit en code. Il s’agit du moyen le plus simple de découvrir comment les différentes fonctionnalités fonctionnent Office scripts. Si vous avez besoin de code pour une action spécifique, basculez vers l’enregistreur d’actions, effectuez les actions, sélectionnez Copier en tant que **code** et collez le code résultant dans votre script.

:::image type="content" source="../images/action-recorder-copy-code.png" alt-text="Volet des tâches de l’enregistreur d’actions avec le bouton « Copier en tant que code » en surbrillance.":::

## <a name="verify-an-object-is-present"></a>Vérifier la présence d’un objet

Les scripts s’appuient souvent sur une feuille de calcul ou une table en cours de présence dans le workbook. Toutefois, ils peuvent être renommés ou supprimés entre les séquences de script. En vérifiant si ces tables ou feuilles de calcul existent avant d’y appeler des méthodes, vous pouvez vous assurer que le script ne se termine pas brusquement.

L’exemple de code suivant vérifie si la feuille de calcul « Index » est présente dans le manuel. Si la feuille de calcul est présente, le script obtient une plage et continue. S’il n’est pas présent, le script enregistre un message d’erreur personnalisé.

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

L’opérateur TypeScript `?` vérifie si l’objet existe avant d’appeler une méthode. Cela peut simplifier votre code si vous n’avez rien de spécial à faire lorsque l’objet n’existe pas.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Valider d’abord les données et l’état du workbook

Assurez-vous que toutes vos feuilles de calcul, tableaux, formes et autres objets sont présents avant de travailler sur les données. À l’aide du modèle précédent, vérifiez si tout se trouve dans le workbook et correspond à vos attentes. Le fait de le faire avant l’écriture de données garantit que votre script ne laisse pas le workbook dans un état partiel.

Le script suivant requiert la présence de deux tables nommées « Table1 » et « Table2 ». Le script vérifie d’abord si les tables `return` sont présentes, puis se termine par l’instruction et un message approprié si ce n’est pas le cas.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue...
}
```

Si la vérification se produit dans une fonction distincte, vous devez quand même mettre fin au script en émettant `return` l’instruction à partir de la `main` fonction. Le retour à partir de la sous-partie ne termine pas le script.

Le script suivant a le même comportement que le précédent. La différence est que la fonction `main` appelle la `inputPresent` fonction pour tout vérifier. `inputPresent` renvoie un booléen (`true` ou `false`) pour indiquer si toutes les entrées requises sont présentes. La `main` fonction utilise ce type booléen pour décider de poursuivre ou de mettre fin au script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue...
}

function inputPresent(workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a>Quand utiliser une `throw` instruction

Une [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) instruction indique qu’une erreur inattendue s’est produite. Il termine immédiatement le code. En grande partie, vous n’avez pas besoin de le `throw` faire à partir de votre script. En règle générale, le script informe automatiquement l’utilisateur que le script n’a pas réussi à s’exécuter en raison d’un problème. Dans la plupart des cas, il suffit de terminer le script avec un message d’erreur et une `return` instruction de la `main` fonction.

Toutefois, si votre script s’exécute dans le cadre d’Power Automate flux, vous voudrez peut-être arrêter le flux de continuer. Une `throw` instruction arrête le script et indique au flux de s’arrêter également.

Le script suivant montre comment utiliser l’instruction `throw` dans notre exemple de vérification de table.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a>Quand utiliser une `try...catch` instruction

L’instruction [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) permet de détecter si un appel d’API échoue et de continuer à l’exécution du script.

Prenons l’extrait de code suivant qui effectue une mise à jour de données importante sur une plage.

```TypeScript
range.setValues(someLargeValues);
```

Si `someLargeValues` la taille est supérieure Excel sur le Web peut gérer, l’appel `setValues()` échoue. Le script échoue également avec une [erreur d’runtime](../testing/troubleshooting.md#runtime-errors). L’instruction `try...catch` permet à votre script de reconnaître cette condition, sans terminer immédiatement le script et afficher l’erreur par défaut.

Une approche pour offrir à l’utilisateur du script une meilleure expérience consiste à lui présenter un message d’erreur personnalisé. L’extrait de code suivant montre une instruction `try...catch` consignant plus d’informations sur les erreurs pour mieux aider le lecteur.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Une autre approche de traitement des erreurs consiste à avoir un comportement de retour qui gère le cas d’erreur. L’extrait de code suivant utilise `catch` le bloc pour essayer une autre méthode décomposer la mise à jour en plus petites parties et éviter l’erreur.

> [!TIP]
> Pour obtenir un exemple complet sur la mise à jour d’une grande plage, voir [Écrire un jeu de données de grande taille](../resources/samples/write-large-dataset.md).

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> L’utilisation `try...catch` à l’intérieur ou autour d’une boucle ralentit votre script. Pour plus d’informations sur les performances, voir [Éviter d’utiliser des `try...catch` blocs](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>Voir aussi

- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Informations de dépannage pour les Power Automate avec Office scripts](../testing/power-automate-troubleshooting.md)
- [Limites de plateforme avec Office scripts](../testing/platform-limits.md)
- [Améliorer les performances de vos scripts Office de gestion](web-client-performance.md)
