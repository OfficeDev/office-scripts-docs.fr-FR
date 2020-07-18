---
title: Principes de base des scripts pour Office Scripts dans Excel sur le web
description: Informations sur le modèle d’objet et autres concepts de base pour vous familiariser avec les scripts Office.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: 6c02f4fb986e6a0ed1dd7afb099aaa1c9d1ea276
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160473"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Principes de base des scripts pour Office Scripts dans Excel sur le web (préversion)

Cet article vous présente les aspects techniques de Office Scripts. Vous découvrirez comment les objets Excel fonctionnent ensemble et comment l’éditeur de code se synchronise avec un classeur.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a>Fonction `main` :

Chaque script Office doit contenir une fonction `main` avec le type `ExcelScript.Workbook` comme premier paramètre. Une fois la fonction exécutée, l’application Excel appelle cette fonction `main` en fournissant le classeur en tant que premier paramètre. Par conséquent, il est important de ne pas modifier la signature de base de la fonction `main` une fois que vous avez enregistré le script ou créé un nouveau script à partir de l’éditeur de code.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

Le code à l’intérieur de la fonction `main` s’exécute lors de l’exécution du script. `main` peut appeler d’autres fonctions dans le script, mais le code qui n’est pas inclus dans une fonction ne s’exécutera pas.

> [!CAUTION]
> Si votre fonction `main` se présente comme `async function main(context: Excel.RequestContext)`, votre script utilise l’ancien modèle API asynchrone. Pour plus d’informations (notamment sur la conversion de votre script vers le modèle API actuel), consultez [Prendre en charge les anciens scripts Office qui utilisent les API asynchrone](excel-async-model.md).

## <a name="object-model"></a>Modèle d’objet

Pour écrire un script, vous devez comprendre la manière dont les API de script Office s’adaptent. Les composants d’un classeur sont dépendants les uns des autres. Dans de nombreux cas, ces relations correspondent à celles de l’interface utilisateur d’Excel.

- Un **classeur** contient une ou plusieurs **feuilles de calcul**.
- Une **feuille de calcul** donne accès à des cellules via **plage** objets.
- Une **plage** représente un groupe de cellules contiguës.
- Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.
- Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.
- Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.

### <a name="workbook"></a>Classeur

Chaque script est fourni avec un `workbook` objet de type `Workbook` par la fonction `main`. Il s’agit de l’objet de niveau supérieur par lequel votre script interagit avec le classeur Excel.

Le script suivant permet d’obtenir le nom de la feuille de calcul active du classeur.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a>Plages

Une plage est un groupe de cellules contiguës dans le classeur. Les scripts utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la colonne **B** et de la ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et des lignes **2** à **4**) pour définir les plages.

Les plages ont trois propriétés principales : valeurs, formules et format. Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules. Ils sont accessibles via `getValues`, `getFormulas`et `getFormat`. Les valeurs et les formules peuvent être modifiées avec `setValues` et `setFormulas`, tandis que le format est un objet `RangeFormat` composé de plusieurs objets de plus petite taille définis individuellement.

Les plages utilisent des tableaux à deux dimensions pour gérer les informations. Pour plus d’informations sur la gestion de ces tableaux dans la structure de scripts Office, consultez la section [utilisation des plages de la section utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md#working-with-ranges).

#### <a name="range-sample"></a>Exemple de plage

L’exemple de code suivant montre comment créer des registres des ventes. Ce script utilise `Range` objets pour déterminer les valeurs, les formules et les parties de la mise en forme.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create the headers and format them to stand out.
    let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
    let headerRange = sheet.getRange("B2:E2");
    headerRange.setValues(headers);
    headerRange.getFormat().getFill().setColor("#4472C4");
    headerRange.getFormat().getFont().setColor("white");

    // Create the product data rows.
    let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.56],
    ];
    let dataRange = sheet.getRange("B3:D5");
    dataRange.setValues(productData);

    // Create the formulas to total the amounts sold.
    let totalFormulas = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"],
    ];
    let totalRange = sheet.getRange("E3:E6");
    totalRange.setFormulas(totalFormulas);
    totalRange.getFormat().getFont().setBold(true);

    // Display the totals as US dollar amounts.
    totalRange.setNumberFormat("$0.00");
}
```

L’exécution de ce script crée les données suivantes dans la feuille de calcul active :

![Un registre des ventes affiche des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Graphiques, tableaux et autres objets de données

Les scripts peuvent créer et manipuler les structures de données et les visualisations dans Excel. Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore. Celles-ci sont stockées dans des collections, qui seront décrites plus loin dans cet article.

#### <a name="creating-a-table"></a>Création d’un tableau

Créez des tableaux à l’aide des plages de données remplies. Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.

L’exemple de code suivant crée un tableau à l’aide des plages de l’exemple précédent.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

L’exécution de ce script sur la feuille de calcul avec les données précédentes crée le tableau suivant :

![Un tableau créée à partir du registre des ventes précédent.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Création d’un graphique

Vous pouvez créer un graphique pour visualiser les données d’une plage. Les scripts permettent des dizaines de variétés de graphiques, chacune pouvant être personnalisée pour répondre à vos besoins.

Le script suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Create a column chart using the data from B3:C5.
    let chart = sheet.addChart(
        ExcelScript.ChartType.columnStacked,
        sheet.getRange("B3:C5")
    );

    // Set the margin of the chart to be 100 pixels from the top of the screen.
    chart.setTop(100);
}
```

L’exécution de ce script sur la feuille de calcul avec le tableau précédent crée le graphique suivant :

![Un histogramme montrant les quantités pour trois des articles présents dans le registre des ventes précédent.](../images/chart-sample.png)

### <a name="collections-and-other-object-relations"></a>Collections et autres relations d’objets

Tout objet enfant est accessible via son objet parent. Par exemple, vous pouvez lire `Worksheets` à partir de l’objet `Workbook`. Il y aura une méthode de `get` liée sur la classe parente (par exemple, `Workbook.getWorksheets()` ou `Workbook.getWorksheet(name)`). `get` les méthodes qui sont singulières renvoient un objet unique et nécessitent un ID ou un nom pour l’objet spécifique (par exemple, le nom d’une feuille de calcul). `get` les méthodes qui permettent de renvoyer l’ensemble de la collection d’objets sous la forme d’une matrice. Si la collection est vide, vous obtenez une matrice vide (`[]`).

Une fois la collection récupérée, vous pouvez utiliser des opérations de tableau régulières, telles que l’acquisition de ses `length` ou utiliser des `for`, `for..of``while` des boucles pour l’itération ou utiliser des méthodes matricielles telles que les `map``forEach`. Vous pouvez également accéder aux objets individuels dans la collection à l’aide de la valeur d’index de tableau. Par exemple, `workbook.getTables()[0]` renvoie la première table de la collection. Pour en savoir plus sur l’utilisation de la fonctionnalité de tableau intégrée avec l’infrastructure de scripts Office, consultez la section [utilisation des collections de utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md#working-with-collections).

Le script suivant extrait toutes les tables du classeur. Il vérifie ensuite que les en-têtes sont affichés, les boutons de filtre sont visibles et le style de tableau est paramétré sur « TableStyleLight1 ».

```typescript
function main(workbook: ExcelScript.Workbook) {
  /* Get table collection */
  const tables = workbook.getTables();
  /* Set table formatting properties */
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

#### <a name="adding-excel-objects-with-a-script"></a>Ajout d’objets Excel à l’aide d’un script

Vous pouvez ajouter des objets document par programme, tels que des tableaux ou des graphiques, en appelant la méthode `add` correspondante disponible sur l’objet parent.

> [!NOTE]
> N’ajoutez pas manuellement des objets aux tableaux de collections. Utilisez les `add` méthodes sur les objets parents par exemple, ajoutez un `Table` à une `Worksheet` avec la méthode `Worksheet.addTable`.

Le script suivant crée un tableau dans Excel sur la première feuille de calcul du classeur. Notez que la table créée est renvoyée par la méthode `addTable`.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in C3:G10.
    let table = sheet.addTable(
      "C3:G10",
       true /* True because the table has headers. */
    );
}
```

## <a name="removing-excel-objects-with-a-script"></a>Suppression d’objets Excel à l’aide d’un script

Pour supprimer un objet, appelez la méthode de `delete` l’objet.

> [!NOTE]
> Comme pour l’ajout d’objets, ne supprimez pas manuellement les objets des tableaux de collections. Utilisez les méthodes `delete` sur les objets de type collection. Par exemple, supprimez un `Table` d’un `Worksheet` à l’aide d' `Table.delete`.

Le script suivant supprime la première feuille de calcul du classeur.

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a>Lectures complémentaires sur le modèle d’objet

La [Documentation de référence de l’API Office Scripts](/javascript/api/office-scripts/overview) est une liste complète des objets utilisés dans Office Scripts. Si vous souhaitez en savoir plus, vous pouvez accéder aux informations sur la classe de votre choix en utilisant la table des matières. Voici quelques pages fréquemment consultées.

- [Graphique](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [Commentaire](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [Range](/javascript/api/office-scripts/excelscript/excelscript.range)
- [RangeFormat](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [Forme](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [Tableau](/javascript/api/office-scripts/excelscript/excelscript.table)
- [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [Feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a>Voir aussi

- [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](../tutorials/excel-tutorial.md)
- [Lire les données d’un classeur avec les scripts Office dans Excel sur le web](../tutorials/excel-read-tutorial.md)
- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
