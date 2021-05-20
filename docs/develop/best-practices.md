---
title: Meilleures pratiques dans Office scripts
description: Comment prévenir les problèmes courants et écrire des scripts Office qui peuvent gérer des entrées ou des données inattendues.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546024"
---
# <a name="best-practices-in-office-scripts"></a>Meilleures pratiques dans Office scripts

Ces modèles et pratiques sont conçus pour aider vos scripts à fonctionner avec succès à chaque fois. Utilisez-les pour éviter les pièges courants lorsque vous commencez à automatiser votre Excel de travail.

## <a name="verify-an-object-is-present"></a>Vérifier la présente d’un objet

Les scripts s’appuient souvent sur une certaine feuille de travail ou table présente dans le cahier de travail. Toutefois, ils peuvent être renommés ou supprimés entre les scripts. En vérifiant si ces tables ou feuilles de travail existent avant d’appeler des méthodes sur eux, vous pouvez vous assurer que le script ne se termine pas brusquement.

L’exemple de code suivant vérifie si la feuille de travail « Index » est présente dans le cahier de travail. Si la feuille de travail est présente, le script obtient une plage et procède. S’il n’est pas présent, le script enregistre un message d’erreur personnalisé.

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

L’opérateur TypeScript `?` vérifie si l’objet existe avant d’appeler une méthode. Cela peut rendre votre code plus rationalisé si vous n’avez pas besoin de faire quelque chose de spécial lorsque l’objet n’existe pas.

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>Valider d’abord les données et l’état du cahier de travail

Assurez-vous que toutes vos feuilles de travail, tables, formes et autres objets sont présents avant de travailler sur les données. En utilisant le modèle précédent, vérifiez si tout est dans le cahier de travail et correspond à vos attentes. Cela avant qu’une donnée ne soit écrite garantit que votre script ne laisse pas le manuel dans un état partiel.

Le script suivant exige la présente deux tableaux nommés « Tableau1 » et « Tableau2 ». Le script vérifie d’abord si les tables sont présentes, puis se termine par `return` l’instruction et un message approprié si elles ne sont pas.

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

  // Continue....
}
```

Si la vérification se déroule dans une fonction distincte, vous devez toujours mettre fin au script en émettant `return` l’instruction à partir de la `main` fonction. Revenir de la sous-fonction ne termine pas le script.

Le script suivant a le même comportement que le précédent. La différence est que la `main` fonction appelle la fonction pour tout `inputPresent` vérifier. `inputPresent` renvoie un boolean `true` (ou `false` ) pour indiquer si toutes les entrées requises sont présentes. La `main` fonction utilise ce boolean pour décider de continuer ou de terminer le script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
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

Une [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) déclaration indique qu’une erreur inattendue s’est produite. Il termine le code immédiatement. Pour la plupart, vous n’avez pas besoin de `throw` de votre script. Habituellement, le script informe automatiquement l’utilisateur que le script n’a pas été exécuté en raison d’un problème. Dans la plupart des cas, il suffit de terminer le script par un message d’erreur et une `return` instruction de la `main` fonction.

Toutefois, si votre script est en cours d’exécution dans le cadre d’Power Automate flux de flux, vous pouvez empêcher le flux de continuer. Une `throw` instruction arrête le script et indique que le flux s’arrête aussi.

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

[`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)L’instruction est un moyen de détecter si un appel API échoue et continuer à exécuter le script.

Considérez l’extrait suivant qui effectue une mise à jour de données importante sur une plage.

```TypeScript
range.setValues(someLargeValues);
```

Si `someLargeValues` elle est plus grande Excel sur le Web peut gérer, l’appel `setValues()` échoue. Le script échoue alors également avec une erreur [de temps d’exécution](../testing/troubleshooting.md#runtime-errors). `try...catch`L’instruction permet à votre script de reconnaître cette condition, sans mettre immédiatement fin au script et afficher l’erreur par défaut.

Une approche pour donner à l’utilisateur de script une meilleure expérience est de leur présenter un message d’erreur personnalisé. L’extrait suivant montre une instruction enregistrant `try...catch` plus d’informations d’erreur pour mieux aider le lecteur.

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

Une autre approche pour traiter les erreurs est d’avoir un comportement de récupération qui gère le cas d’erreur. L’extrait suivant utilise le bloc `catch` pour essayer une autre méthode briser la mise à jour en petits morceaux et éviter l’erreur.

> [!TIP]
> Pour un exemple complet sur la façon de mettre à jour une large plage, voir [Ecrire un grand ensemble de données](../resources/samples/write-large-dataset.md).

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
> `try...catch`L’utilisation à l’intérieur ou autour d’une boucle ralentit votre script. Pour plus d’informations sur les performances, voir [Éviter d’utiliser `try...catch` des blocs](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).

## <a name="see-also"></a>Voir aussi

- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Informations de dépannage pour Power Automate avec Office scripts](../testing/power-automate-troubleshooting.md)
- [Limites de plate-forme avec Office scripts](../testing/platform-limits.md)
- [Améliorez les performances de vos scripts Office’argent](web-client-performance.md)
