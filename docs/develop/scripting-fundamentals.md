---
title: Principes de base des scripts pour Office Scripts dans Excel sur le web
description: Informations sur le modèle d’objet et autres concepts de base pour vous familiariser avec les scripts Office.
ms.date: 04/24/2020
localization_priority: Priority
ms.openlocfilehash: 8449654e359f665677f3d416a8e28fa4d6930f26
ms.sourcegitcommit: 350bd2447f616fa87bb23ac826c7731fb813986b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2020
ms.locfileid: "43919797"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Principes de base des scripts pour Office Scripts dans Excel sur le web (préversion)

Cet article vous présente les aspects techniques de Office Scripts. Vous découvrirez comment les objets Excel fonctionnent ensemble et comment l’éditeur de code se synchronise avec un classeur.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Modèle d’objet

Pour comprendre les API Excel, vous devez connaître la manière dont les composants d’un classeur sont liés les uns aux autres.

- Un **classeur** contient une ou plusieurs **feuilles de calcul**.
- Une **feuille de calcul** donne accès à des cellules via **plage** objets.
- Une **plage** représente un groupe de cellules contiguës.
- Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.
- Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.
- Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.

### <a name="ranges"></a>Plages

Une plage est un groupe de cellules contiguës dans le classeur. Les scripts utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la colonne **B** et de la ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et des lignes **2** à **4**) pour définir les plages.

Les plages comportent trois propriétés principales : `values`, `formulas`et `format`. Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules.

#### <a name="range-sample"></a>Exemple de plage

L’exemple de code suivant montre comment créer des registres des ventes. Le script utilise les objets `Range` pour déterminer les valeurs, les formules et les formats.

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

![Un registre des ventes affiche des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Graphiques, tableaux et autres objets de données

Les scripts peuvent créer et manipuler les structures de données et les visualisations dans Excel. Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore.

#### <a name="creating-a-table"></a>Création d’un tableau

Créez des tableaux à l’aide des plages de données remplies. Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.

L’exemple de code suivant crée un tableau à l’aide des plages de l’exemple précédent.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

L’exécution de ce script sur la feuille de calcul avec les données précédentes crée le tableau suivant :

![Un tableau créée à partir du registre des ventes précédent.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Création d’un graphique

Vous pouvez créer un graphique pour visualiser les données d’une plage. Les scripts permettent des dizaines de variétés de graphiques, chacune pouvant être personnalisée pour répondre à vos besoins.

Le script suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

L’exécution de ce script sur la feuille de calcul avec le tableau précédent crée le graphique suivant :

![Un histogramme montrant les quantités pour trois des articles présents dans le registre des ventes précédent.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Lectures complémentaires sur le modèle d’objet

La [Documentation de référence de l’API Office Scripts](/javascript/api/office-scripts/overview) est une liste complète des objets utilisés dans Office Scripts. Si vous souhaitez en savoir plus, vous pouvez accéder aux informations sur la classe de votre choix en utilisant la table des matières. Voici quelques pages fréquemment consultées.

- [Graphique](/javascript/api/office-scripts/excel/excel.chart)
- [Commentaire](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Forme](/javascript/api/office-scripts/excel/excel.shape)
- [Tableau](/javascript/api/office-scripts/excel/excel.table)
- [Classeur](/javascript/api/office-scripts/excel/excel.workbook)
- [Feuille de calcul](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>Fonction `main` :

Chaque script Office doit contenir une fonction `main` avec la signature suivante, qui inclut la définition de type `Excel.RequestContext` :

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

Le code à l’intérieur de la fonction `main` s’exécute lors de l’exécution du script. `main` peut appeler d’autres fonctions dans le script, mais le code qui n’est pas inclus dans une fonction ne s’exécutera pas.

## <a name="context"></a>Contexte

La fonction `main` accepte un paramètre `Excel.RequestContext`, nommé `context`. Vous devez imaginer le `context` comme un pont entre le script et le classeur. Le script accède au classeur avec l’objet `context` et utilise ce `context` pour envoyer et recevoir des données.

L’objet `context` est nécessaire car le script et Excel sont exécutés dans différents processus et emplacements. Le script doit apporter des modifications ou rechercher les données du classeur dans le cloud. L’objet `context` gère ces opérations.

## <a name="sync-and-load"></a>Synchronisation et chargement

Comme le script et le classeur s’exécutent dans des emplacements différents, le transfert de données entre les deux prend du temps. Pour améliorer les performances du script, les commandes sont mises en file d’attente jusqu’à ce que le script appelle explicitement l’opération `sync` pour synchroniser le script et le classeur. Le script peut fonctionner de façon indépendante jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :

- Lisez les données du classeur (en suivant une `load`opération de ou une méthode qui renvoie une [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult)).
- Écrire les données dans le classeur (généralement quand le script est terminé).

L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :

![Un diagramme montrant les opérations de lecture et d’écriture effectuées dans le classeur à partir du script.](../images/load-sync.png)

### <a name="sync"></a>Synchronisation

Lorsque le script a besoin de lire ou d’écrire des données dans le classeur, appelez la méthode `RequestContext.sync` comme illustré ici :

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()` est appelé implicitement à la fin d’un script.

Une fois l’opération `sync` terminée, le classeur se met à jour pour illustrer les opérations d’écriture que le script a spécifiées. Une opération d’écriture définit une propriété sur un objet Excel (par exemple : `range.format.fill.color = "red"`) ou appelle une méthode qui modifie une propriété (par exemple : `range.format.autoFitColumns()`). L’opération `sync` lit également les valeurs du classeur demandées par le script à l’aide d’une opération `load` ou d’une méthode renvoyant une `ClientResult` (comme indiqué dans la section suivante).

La synchronisation du script avec le classeur peut prendre du temps, en fonction de votre réseau. Vous devez diminuer le nombre d’appels `sync` pour faciliter l’exécution du script.  

### <a name="load"></a>Charger

Un script doit charger les données du classeur avant de les lire. Toutefois, le chargement fréquent de données à partir de l’intégralité du classeur réduirait considérablement la vitesse du script. La méthode `load`, qui permet au script d’indiquer spécifiquement les données du classeur à récupérer, est plus appropriée.

La méthode `load` est disponible sur tous les objets Excel. Le script doit charger les propriétés d’un objet avant de pouvoir les lire. Sinon, cela entraînera une erreur.

Les exemples suivants utilisent un objet `Range` pour illustrer les trois méthodes utilisées par `load` pour charger les données.

|Objectif |Exemple de commande | Effet |
|:--|:--|:--|
|Charger une propriété |`myRange.load("values");` | Charge une seule propriété. Dans ce cas, le tableau à deux dimensions des valeurs dans cette plage. |
|Charger plusieurs propriétés |`myRange.load("values, rowCount, columnCount");`| Charge toutes les propriétés d’une liste, qui sont délimitées par des virgules. Dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes. |
|Tout charger | `myRange.load();`|Charge toutes les propriétés de la plage. Ceci n’est pas une solution recommandée, car elle ralentit le script, qui charge des données superflues. Vous devez utiliser cette opération uniquement lorsque vous testez le script ou si vous avez besoin de toutes les propriétés de l’objet. |

Le script doit appeler `context.sync()` avant de lire les valeurs chargées.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

Vous pouvez également charger des propriétés sur l’ensemble d’une collection. Chaque objet d’une collection possède une propriété `items`, qui est un tableau contenant les objets dans cette collection. L’utilisation de `items` comme point de départ d’un appel hiérarchique (`items\myProperty`) pour que `load` charge les propriétés spécifiées sur chacun de ces éléments. L’exemple suivant charge la propriété `resolved` sur tous les objets `Comment` dans l’objet `CommentCollection` d’une feuille de calcul.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Si vous souhaitez en savoir plus sur l’utilisation des collections dans les scripts Office, consultez l’article [Section du tableau sur l'utilisation d'objets JavaScript intégrés dans Office Scripts](javascript-objects.md#array).

### <a name="clientresult"></a>ClientResult

Les méthodes qui renvoient des informations du classeur présentent un modèle semblable au paradigme `load`/`sync`. Par exemple, `TableCollection.getCount` obtient le nombre de tableaux dans la collection. `getCount` renvoie une `ClientResult<number>`, ce qui signifie que la propriété `value` dans le renvoie `ClientResult` est un nombre. Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé. À l’instar du chargement d’une propriété, la valeur `value` est une valeur « vide » locale jusqu’à cet appel`sync`.

Le script suivant fournit le nombre total de tableaux dans le classeur et enregistre ce nombre sur la console.

```TypeScript
async function main(context: Excel.RequestContext) {
  let tableCount = context.workbook.tables.getCount();

  // This sync call implicitly loads tableCount.value.
  // Any other ClientResult values are loaded too.
  await context.sync();

  // Trying to log the value before calling sync would throw an error.
  console.log(tableCount.value);
}
```

## <a name="see-also"></a>Voir aussi

- [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](../tutorials/excel-tutorial.md)
- [Lire les données d’un classeur avec les scripts Office dans Excel sur le web](../tutorials/excel-read-tutorial.md)
- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
