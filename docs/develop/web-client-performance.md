---
title: Améliorer les performances de vos scripts Office de gestion
description: Créez des scripts plus rapides en comprenant la communication entre le Excel et votre script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: a5bd879625b9c3bac0caa621dde312f7c961dd5c
ms.sourcegitcommit: 2aaf7dc527cb6c9f1206550b2c5745280503b2a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/16/2021
ms.locfileid: "52957699"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Améliorer les performances de vos scripts Office de gestion

L’objectif de Office Scripts est d’automatiser une série de tâches couramment exécutées pour vous faire gagner du temps. Un script lent peut avoir l’impression qu’il n’accélère pas votre flux de travail. La plupart du temps, votre script sera parfaitement correct et s’exécutera comme prévu. Toutefois, il existe quelques scénarios qui peuvent avoir une incidence sur les performances.

La raison la plus courante d’un script lent est une communication excessive avec le workbook. Votre script s’exécute sur votre ordinateur local, tandis que le workbook existe dans le cloud. À certains moments, votre script synchronise ses données locales avec celle du workbook. Cela signifie que toutes les opérations d’écriture (telles que ) ne sont appliquées aubook que lorsque cette synchronisation en `workbook.addWorksheet()` arrière-plan se produit. De même, toutes les opérations de lecture (telles que ) obtiennent uniquement des données du manuel pour le `myRange.getValues()` script à ce moment-là. Dans les deux cas, le script récupère des informations avant d’agir sur les données. Par exemple, le code suivant enregistre avec précision le nombre de lignes dans la plage utilisée.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Les API de scripts garantissent que toutes les données du workbook ou du script sont précises et à jour si nécessaire. Vous n’avez pas besoin de vous soucier de ces synchronisations pour que votre script s’exécute correctement. Toutefois, une connaissance de cette communication entre les scripts et le cloud peut vous aider à éviter les appels réseau inutiles.

## <a name="performance-optimizations"></a>Optimisation des performances

Vous pouvez appliquer des techniques simples pour réduire la communication vers le cloud. Les modèles suivants permettent d’accélérer vos scripts.

- Lire les données du workbook une seule fois plutôt que de manière répétée dans une boucle.
- Supprimez les `console.log` instructions inutiles.
- Évitez d’utiliser des blocs try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Lire les données d’un workbook en dehors d’une boucle

Toute méthode qui obtient des données à partir du manuel peut déclencher un appel réseau. Au lieu d’effectuer le même appel à plusieurs reprises, vous devez enregistrer les données localement chaque fois que cela est possible. Cela est particulièrement vrai lorsque vous traitez des boucles.

Envisagez un script pour obtenir le nombre de nombres négatifs dans la plage utilisée d’une feuille de calcul. Le script doit itérer sur chaque cellule de la plage utilisée. Pour ce faire, il a besoin de la plage, du nombre de lignes et du nombre de colonnes. Vous devez les stocker en tant que variables locales avant de démarrer la boucle. Dans le cas contraire, chaque itération de la boucle force un retour au workbook.

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
> En tant qu’expérience, essayez de remplacer `usedRangeValues` dans la boucle par `usedRange.getValues()` . Vous remarquerez peut-être que l’exécuter du script prend beaucoup plus de temps lorsque vous traitez des plages importantes.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Évitez `try...catch` d’utiliser des blocs dans des boucles dans ou autour

Nous vous déconseillons d’utiliser des instructions en [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) boucle ou autour de boucles. C’est pour la même raison que vous devez éviter la lecture de données dans une boucle : chaque itération force le script à se synchroniser avec le workbook pour s’assurer qu’aucune erreur n’a été lancée. La plupart des erreurs peuvent être évitées en vérifiant les objets renvoyés à partir du workbook. Par exemple, le script suivant vérifie que la table renvoyée par le workbook existe avant d’essayer d’ajouter une ligne.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Supprimer les `console.log` instructions inutiles

La journalisation de la console est un outil vital [pour le débogage de vos scripts.](../testing/troubleshooting.md) Toutefois, il force le script à se synchroniser avec le workbook pour s’assurer que les informations consignées sont à jour. Envisagez de supprimer les instructions de journalisation inutiles (telles que celles utilisées pour les tests) avant de partager votre script. En règle générale, cela ne provoque pas de problème de performances perceptible, sauf si `console.log()` l’instruction est en boucle.

## <a name="case-by-case-help"></a>Aide au cas par cas

À mesure que la plateforme Office Scripts s’étend [](/adaptive-cards)pour fonctionner avec [Power Automate,](https://flow.microsoft.com/)les cartes adaptatives et d’autres fonctionnalités entre produits, les détails de la communication de script-workbook deviennent plus complexes. Si vous avez besoin d’aide pour accélérer l’exécuter, [contactez-vous](/answers/topics/office-scripts-excel-dev.html)via Microsoft Q&A . N’oubliez pas de baliser votre question avec « office-scripts-dev » afin que les experts la trouvent et vous aident.

## <a name="see-also"></a>Voir aussi

- [Principes de base pour la rédaction de scripts Office en Excel sur le web](scripting-fundamentals.md)
- [Documentation web MDN : boucles et itération](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
