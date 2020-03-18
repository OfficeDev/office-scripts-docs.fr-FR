---
title: Exemples de scripts pour les scripts Office dans Excel sur le Web
description: Collection d’exemples de code à utiliser avec des scripts Office dans Excel sur le Web.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: abb4064dfde8b644035e725832e481e6463e979e
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700244"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a>Exemples de scripts pour les scripts Office dans Excel sur le Web (aperçu)

Les exemples suivants sont des scripts simples que vous pouvez essayer dans vos propres classeurs. Pour les utiliser dans Excel sur le Web :

1. Ouvrir l’onglet **automatiser** .
2. Appuyez sur **éditeur de code**.
3. Appuyez sur **nouveau script** dans le volet Office de l’éditeur de code.
4. Remplacez l’intégralité du script par l’exemple de votre choix.
5. Appuyez sur **exécuter** dans le volet Office de l’éditeur de code.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a>Concepts de base des scripts

Ces exemples illustrent des blocs de construction fondamentaux pour les scripts Office. Ajoutez-les à vos scripts pour étendre votre solution et résoudre les problèmes courants.

### <a name="read-and-log-one-cell"></a>Lecture et journalisation d’une cellule

Cet exemple lit la valeur de **a1** et l’imprime sur la console.

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a>Utiliser des dates

Cet exemple utilise l’objet JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) pour obtenir la date et l’heure actuelles, puis écrit ces valeurs dans deux cellules de la feuille de calcul active.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

## <a name="display-data"></a>Afficher les données

Ces exemples montrent comment utiliser les données de feuille de calcul et fournir aux utilisateurs une meilleure vue ou organisation.

### <a name="apply-conditional-formatting"></a>Application d’une mise en forme conditionnelle

Cet exemple applique la mise en forme conditionnelle à la plage utilisée dans la feuille de calcul. La mise en forme conditionnelle est un remplissage vert pour les 10% de valeurs les plus fréquentes.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a>Créer un tableau trié

Cet exemple montre comment créer un tableau à partir de la plage utilisée dans la feuille de calcul active, puis comment le trier en fonction de la première colonne.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a>Collaboration

Ces exemples montrent comment utiliser les fonctionnalités liées à la collaboration d’Excel, telles que les commentaires.

### <a name="delete-resolved-comments"></a>Supprimer les commentaires résolus

Cet exemple montre comment supprimer tous les commentaires résolus de la feuille de calcul active.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a>Exemples de scénario

Pour obtenir des exemples illustrant des solutions plus étendues dans le monde réel, consultez [exemples de scénarios pour les scripts Office](scenarios/sample-scenario-overview.md).

## <a name="suggest-new-samples"></a>Suggérer de nouveaux exemples

Nous vous invitons à suggérer de nouveaux exemples. S’il existe un scénario courant qui aide les autres développeurs de script, veuillez nous en indiquer dans la section commentaires ci-dessous.