---
title: Améliorez les performances de vos scripts Office’argent
description: Créez des scripts plus rapides en comprenant la communication entre Excel manuel et votre script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544990"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Améliorez les performances de vos scripts Office’argent

Le but de Office scripts est d’automatiser des séries de tâches couramment exécutées pour vous faire gagner du temps. Un script lent peut avoir l’impression qu’il n’accélère pas votre flux de travail. La plupart du temps, votre script sera parfaitement bien et exécuté comme prévu. Cependant, il existe quelques scénarios évitables qui peuvent affecter les performances.

La raison la plus courante d’un script lent est une communication excessive avec le cahier de travail. Votre script s’exécute sur votre machine locale, tandis que le manuel existe dans le cloud. À certains moments, votre script synchronise ses données locales avec celle du cahier de travail. Cela signifie que toutes les opérations d’écriture (telles `workbook.addWorksheet()` que) ne sont appliquées au manuel que lorsque cette synchronisation en coulisses se produit. De même, toutes les opérations de lecture (telles `myRange.getValues()` que) ne proviennent que des données du cahier de travail pour le script à ces moments. Dans les deux cas, le script récupère des informations avant qu’elles n’agissent sur les données. Par exemple, le code suivant enregistrera avec précision le nombre de lignes dans la plage utilisée.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office Les API scripts garantissent que toutes les données du cahier de travail ou du script sont exactes et à jour si nécessaire. Vous n’avez pas besoin de vous soucier de ces synchronisations pour que votre script s’exécute correctement. Toutefois, une prise de conscience de cette communication script-cloud peut vous aider à éviter les appels réseau non désaillés.

## <a name="performance-optimizations"></a>Optimisations de performance

Vous pouvez appliquer des techniques simples pour aider à réduire la communication vers le cloud. Les modèles suivants aident à accélérer vos scripts.

- Lisez les données du manuel une fois au lieu de plusieurs fois en boucle.
- Supprimez les `console.log` instructions inutiles.
- Évitez d’utiliser des blocs try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Lire les données du cahier de travail en dehors d’une boucle

Toute méthode qui obtient des données du cahier de travail peut déclencher un appel réseau. Plutôt que de faire à plusieurs reprises le même appel, vous devez enregistrer des données localement dans la mesure du possible. Cela est particulièrement vrai lorsqu’il s’agit de boucles.

Considérez un script pour obtenir le nombre de nombres négatifs dans la plage utilisée d’une feuille de travail. Le script doit itérer sur chaque cellule de la plage utilisée. Pour ce faire, il a besoin de la plage, le nombre de lignes, et le nombre de colonnes. Vous devez les stocker comme variables locales avant de commencer la boucle. Dans le cas contraire, chaque itération de la boucle forcera un retour au cahier de travail.

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
> Comme une expérience, essayez de remplacer `usedRangeValues` dans la boucle par `usedRange.getValues()` . Vous remarquerez peut-être que le script prend beaucoup plus de temps à exécuter lorsqu’il s’agit de grandes plages.

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>Évitez d’utiliser `try...catch` des blocs dans ou autour des boucles

Nous ne recommandons pas d’utiliser les [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instructions en boucle ou en boucles environnantes. C’est pour la même raison que vous devez éviter de lire les données en boucle : chaque itération oblige le script à se synchroniser avec le cahier de travail pour s’assurer qu’aucune erreur n’a été lancée. La plupart des erreurs peuvent être évitées en vérifiant les objets retournés du cahier de travail. Par exemple, le script suivant vérifie que la table retournée par le cahier de travail existe avant d’essayer d’ajouter une ligne.

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

### <a name="remove-unnecessary-consolelog-statements"></a>Supprimer les instructions `console.log` inutiles

L’enregistrement des consoles est un outil essentiel [pour débouger vos scripts.](../testing/troubleshooting.md) Toutefois, il oblige le script à se synchroniser avec le cahier de travail pour s’assurer que les informations enregistrées sont à jour. Envisagez de supprimer les instructions d’enregistrement inutiles (telles que celles utilisées pour les tests) avant de partager votre script. Cela ne cause généralement pas de problème de performances notables, sauf si `console.log()` l’instruction est en boucle.

## <a name="case-by-case-help"></a>Aide au cas par cas

Au fur et à mesure que la plate-forme Office Scripts [s’étend pour fonctionner avec Power Automate,](https://flow.microsoft.com/)Adaptive [Cards](/adaptive-cards)et d’autres fonctionnalités de produits croisés, les détails de la communication script-cahier de travail deviennent plus complexes. Si vous avez besoin d’aide pour faire fonctionner votre script plus rapidement, s’il vous plaît [tendre la main via Microsoft Q&A](/answers/topics/office-scripts-dev.html). Assurez-vous d’étiqueter votre question avec « office-scripts-dev » afin que les experts puissent la trouver et vous aider.

## <a name="see-also"></a>Voir aussi

- [Principes de base pour la rédaction de scripts Office en Excel sur le web](scripting-fundamentals.md)
- [MDN web docs: Boucles et itération](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
