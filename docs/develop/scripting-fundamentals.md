---
title: Scripts de base pour les scripts Office dans Excel sur le Web
description: Informations sur le modèle objet et autres notions de base à connaître avant d’écrire des scripts Office.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700200"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Notions de base sur les scripts pour les scripts Office dans Excel sur le Web (aperçu)

Cet article vous présente les aspects techniques des scripts Office. Vous découvrirez comment les objets Excel fonctionnent ensemble et comment l’éditeur de code se synchronise avec un classeur.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Modèle d’objet

Pour comprendre les API Excel, vous devez comprendre comment les composants d’un classeur sont liés les uns aux autres.

- Un **classeur** contient une ou plusieurs **feuilles de calcul**.
- Une **feuille de calcul** donne accès aux cellules par le biais d’objets **Range** .
- Une **plage** représente un groupe de cellules contiguës.
- Les **plages** permettent de créer et de placer des **tableaux**, des **graphiques**, des **formes**et d’autres objets d’organisation ou de visualisation de données.
- Une **feuille de calcul** contient des collections de ces objets de données qui sont présents dans la feuille individuelle.
- Les **classeurs** contiennent des collections de certains de ces objets de données (tels que les **tableaux**) pour l’intégralité du **classeur**.

### <a name="ranges"></a>Ranges

Une plage est un groupe de cellules contiguës dans le classeur. Les scripts utilisent généralement la notation de style a1 (par exemple, **B3** pour la cellule unique de la ligne **B** et de la colonne **3** ou **C2 : F4** pour les cellules des lignes **C** à **F** et les colonnes **2** à **4**) pour définir des plages.

Les plages ont trois propriétés principales `values`: `formulas`, et `format`. Ces propriétés obtiennent ou définissent les valeurs de cellule, les formules à évaluer et la mise en forme visuelle des cellules.

#### <a name="range-sample"></a>Exemple de plage

L’exemple suivant montre comment créer des enregistrements de ventes. Ce script utilise `Range` des objets pour définir les valeurs, les formules et les formats.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

L’exécution de ce script crée les données suivantes dans la feuille de calcul active :

![Un enregistrement de ventes affichant des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Graphiques, tableaux et autres objets de données

Les scripts peuvent créer et manipuler les structures de données et les visualisations dans Excel. Les tableaux et les graphiques sont deux des objets les plus couramment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et plus encore.

#### <a name="creating-a-table"></a>Création d’un tableau

Créer des tables à l’aide de plages de données. La mise en forme et les contrôles de tableau (tels que les filtres) sont automatiquement appliqués à la plage.

Le script suivant crée un tableau à l’aide des plages de l’exemple précédent.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

L’exécution de ce script sur la feuille de calcul avec les données précédentes crée le tableau suivant :

![Tableau créé à partir de l’enregistrement de ventes précédent.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Création d’un graphique

Créez des graphiques pour visualiser les données dans une plage. Les scripts autorisent des dizaines de variétés de graphiques, chacune pouvant être personnalisée en fonction de vos besoins.

Le script suivant crée un graphique en histogramme simple pour trois éléments et place celle-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

L’exécution de ce script sur la feuille de calcul avec le tableau précédent crée le graphique suivant :

![Histogramme affichant les quantités de trois éléments de l’enregistrement des ventes précédent.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Documentation supplémentaire sur le modèle objet

La [documentation de référence de l’API des scripts Office](/javascript/api/office-scripts/overview) est une liste complète des objets utilisés dans les scripts Office. Vous pouvez utiliser la table des matières pour accéder à une classe sur laquelle vous aimeriez obtenir des informations supplémentaires. Voici quelques-unes des pages fréquemment consultées.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main`elle

Chaque script Office doit contenir une `main` fonction avec la signature suivante, y compris `Excel.RequestContext` la définition de type :

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

Le code à l' `main` intérieur de la fonction s’exécute lors de l’exécution du script. `main`peut appeler d’autres fonctions dans votre script, mais le code qui n’est pas contenu dans une fonction ne s’exécutera pas.

## <a name="context"></a>Contexte

La `main` fonction accepte un `Excel.RequestContext` paramètre, nommé `context`. Imaginez `context` comme un pont entre votre script et le classeur. Votre script accède au classeur avec l' `context` objet et l’utilise `context` pour envoyer et recevoir des données.

L' `context` objet est nécessaire, car le script et Excel s’exécutent dans différents processus et emplacements. Le script devra modifier ou interroger les données du classeur dans le Cloud. L' `context` objet gère ces transactions.

## <a name="sync-and-load"></a>Synchronisation et chargement

Étant donné que votre script et votre classeur s’exécutent à différents emplacements, le transfert de données entre les deux prend du temps. Pour améliorer les performances des scripts, les commandes sont mises en file d’attente jusqu' `sync` à ce que le script appelle explicitement l’opération pour synchroniser le script et le classeur. Votre script peut fonctionner indépendamment jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :

- Lire les données du classeur (après `load` une opération).
- Écrire des données dans le classeur (généralement parce que le script est terminé).

L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :

![Diagramme illustrant les opérations de lecture et d’écriture dans le classeur à partir du script.](../images/load-sync.png)

### <a name="sync"></a>Synchronisation

Chaque fois que votre script a besoin de lire ou d’écrire des données dans le classeur, appelez la `RequestContext.sync` méthode comme illustré ci-dessous :

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()`est implicitement appelé à la fin d’un script.

Une fois `sync` l’opération terminée, le classeur est mis à jour pour refléter toutes les opérations d’écriture que le script a spécifiées. Une opération d’écriture définit une propriété sur un objet Excel (par exemple `range.format.fill.color = "red"`,) ou un appel à une méthode qui modifie une propriété ( `range.format.autoFitColumns()`par exemple,). L' `sync` opération lit également toutes les valeurs du classeur que le script a demandées `load` à l’aide d’une opération (comme indiqué dans la section suivante).

La synchronisation de votre script avec le classeur peut prendre du temps, en fonction de votre réseau. Vous devez réduire le nombre d' `sync` appels pour aider votre script à s’exécuter rapidement.  

### <a name="load"></a>Load

Un script doit charger les données du classeur avant de le lire. Toutefois, le chargement fréquent de données à partir de l’intégralité du classeur réduirait considérablement la vitesse du script. Au lieu de `load` cela, la méthode permet à votre script d’indiquer spécifiquement quelles données doivent être récupérées à partir du classeur.

La `load` méthode est disponible sur tous les objets Excel. Votre script doit charger les propriétés d’un objet avant de pouvoir les lire. Si vous ne le faites pas, une erreur se produit.

Les exemples suivants utilisent un `Range` objet pour afficher les trois façons dont `load` la méthode peut être utilisée pour charger des données.

|Intent |Exemple de commande | Effet |
|:--|:--|:--|
|Charger une propriété |`myRange.load("values");` | Charge une seule propriété, dans ce cas, le tableau à deux dimensions des valeurs de cette plage. |
|Charger plusieurs propriétés |`myRange.load("values, rowCount, columnCount");`| Charge toutes les propriétés à partir d’une liste délimitée par des virgules, dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes. |
|Tout charger | `myRange.load();`|Charge toutes les propriétés de la plage. Il ne s’agit pas d’une solution recommandée, car elle ralentit votre script en obtenant des données inutiles. Vous ne devez l’utiliser que si vous testez votre script ou si vous avez besoin de chaque propriété de l’objet. |

Votre script doit appeler `context.sync()` avant de lire les valeurs chargées.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

Vous pouvez également charger les propriétés sur toute une collection. Chaque objet collection possède une `items` propriété qui est un tableau contenant les objets de cette collection. À `items` l’aide du début d’un appel hiérarchique (`items\myProperty`) pour `load` charger les propriétés spécifiées sur chacun de ces éléments. L’exemple suivant charge la `resolved` propriété sur chaque `Comment` objet dans l' `CommentCollection` objet d’une feuille de calcul.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Pour en savoir plus sur l’utilisation des collections dans les scripts Office, reportez-vous à la [section tableau de l’article Using Built-in JavaScript Objects in Office scripts](javascript-objects.md#array) .

## <a name="see-also"></a>Voir aussi

- [Enregistrer, modifier et créer des scripts Office dans Excel sur le Web](../tutorials/excel-tutorial.md)
- [Lire des données de classeur avec des scripts Office dans Excel sur le Web](../tutorials/excel-read-tutorial.md)
- [Référence de l’API des scripts Office](/javascript/api/office-scripts/overview)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
