---
title: Améliorer les performances de vos scripts Office
description: Créez des scripts plus rapides en vous familiarisant avec la communication entre le classeur Excel et votre script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878805"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Améliorer les performances de vos scripts Office

L’objectif des scripts Office est d’automatiser la série de tâches couramment exécutées pour vous permettre de gagner du temps. Un script lent peut sembler n’accélérer pas votre flux de travail. La plupart du temps, votre script sera parfait et s’exécutera comme prévu. Toutefois, il existe quelques scénarios évitables qui peuvent affecter les performances.

La cause la plus fréquente d’un script lent est une communication excessive avec le classeur. Votre script s’exécute sur votre ordinateur local, tandis que le classeur existe dans le Cloud. À certains moments, votre script synchronise ses données locales avec celles du classeur. Cela signifie que toutes les opérations d’écriture (telles que `workbook.addWorksheet()` ) sont appliquées au classeur uniquement lorsque cette synchronisation en arrière-plan se produit. De même, toutes les opérations de lecture (telles que `myRange.getValues()` ) obtiennent uniquement les données du classeur pour le script à ces moments. Dans les deux cas, le script récupère les informations avant qu’il agisse sur les données. Par exemple, le code suivant consigne exactement le nombre de lignes dans la plage utilisée.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

API de scripts Office Assurez-vous que toutes les données du classeur ou du script sont exactes et à jour, le cas échéant. Vous n’avez pas à vous soucier de ces synchronisations pour que votre script s’exécute correctement. Toutefois, une connaissance de cette communication de script vers le Cloud peut vous aider à éviter les appels réseau inutiles.

## <a name="performance-optimizations"></a>Optimisation des performances

Vous pouvez appliquer des techniques simples pour réduire la communication vers le Cloud. Les modèles suivants permettent d’accélérer vos scripts.

- Lire les données du classeur une seule fois au lieu de répéter dans une boucle.
- Supprimez les `console.log` instructions inutiles.
- Évitez d’utiliser des blocs try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Lire les données du classeur en dehors d’une boucle

Toute méthode qui obtient les données du classeur peut déclencher un appel réseau. Au lieu de faire le même appel de manière répétée, vous devez enregistrer les données localement chaque fois que cela est possible. Cela est particulièrement vrai pour le traitement des boucles.

Considérez un script pour obtenir le nombre de nombres négatifs dans la plage utilisée d’une feuille de calcul. Le script doit parcourir toutes les cellules de la plage utilisée. Pour ce faire, il a besoin de la plage, du nombre de lignes et du nombre de colonnes. Vous devez les stocker en tant que variables locales avant de lancer la boucle. Dans le cas contraire, chaque itération de la boucle forcera un retour au classeur.

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> À titre d’expérimentation, essayez `usedRangeValues` de remplacer dans la boucle par `usedRange.getValues()` . Vous pouvez remarquer que l’exécution du script est beaucoup plus longue lorsque vous traitez des grandes plages.

### <a name="remove-unnecessary-consolelog-statements"></a>Supprimer les `console.log` instructions inutiles

La journalisation de console est un outil essentiel pour [le débogage de vos scripts](../testing/troubleshooting.md). Toutefois, il force le script à se synchroniser avec le classeur afin de s’assurer que les informations consignées sont à jour. Envisagez de supprimer les instructions de journalisation inutiles (telles que celles utilisées pour les tests) avant de partager votre script. Cela ne provoque généralement pas de problèmes de performances perceptibles, sauf si l' `console.log()` instruction est en boucle.

### <a name="avoid-using-trycatch-blocks"></a>Éviter d’utiliser des blocs try/catch

Nous vous déconseillons d’utiliser des [ `try` / `catch` blocs](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) dans le cadre du flux de contrôle attendu d’un script. La plupart des erreurs peuvent être évitées en vérifiant les objets renvoyés à partir du classeur. Par exemple, le script suivant vérifie que la table renvoyée par le classeur existe avant d’essayer d’ajouter une ligne.

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a>Aide cas par cas

À mesure que la plateforme de scripts Office s’étend pour fonctionner avec [Power automate](https://flow.microsoft.com/), [cartes adaptatives](https://docs.microsoft.com/adaptive-cards)et autres fonctionnalités de produit, les détails de la communication de classeur de script deviennent plus compliqués. Si vous avez besoin d’aide pour faire en sorte que votre script s’exécute plus rapidement, veuillez contacter le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts). N’oubliez pas de baliser votre question avec « Office-script » afin que les experts puissent y trouver des rubriques et de l’aide.

## <a name="see-also"></a>Voir aussi

- [Principes de base des scripts pour Office Scripts dans Excel sur le web](scripting-fundamentals.md)
- [NOTIFICATION Web docs : boucles et itération](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
