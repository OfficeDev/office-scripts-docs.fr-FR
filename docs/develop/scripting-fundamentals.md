---
title: Principes de base des scripts pour Office Scripts dans Excel sur le web
description: Informations sur le modèle d’objet et autres concepts de base pour vous familiariser avec les scripts Office.
ms.date: 05/24/2021
ms.localizationpriority: high
ms.openlocfilehash: e2ba7eaa956f2009c9017bbfd1f390f56eb9008e
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585722"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a>Principes de base des scripts pour Scripts Office dans Excel sur le web

Cet article vous présente les aspects techniques de Office Scripts. Vous découvrirez comment les objets Excel fonctionnent ensemble et comment l’éditeur de code se synchronise avec un classeur.

## <a name="typescript-the-language-of-office-scripts"></a>TypeScript : langue des scripts Office

Les scripts Office sont écrits dans [TypeScript](https://www.typescriptlang.org/docs/home.html), qui est un ensemble de scripts [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). Si vous avez l’expérience JavaScript, vos connaissances seront transférées, car la plupart du code est identique dans les deux langages. Nous vous recommandons d'avoir des connaissances en programmation de niveau débutant avant de vous lancer dans le codage de scripts Office. Les ressources suivantes peuvent vous aider à comprendre l'aspect codage des scripts Office.

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a>Fonction `main` : point de départ du script

Chaque script doit contenir une fonction `main` avec le type `ExcelScript.Workbook` comme premier paramètre. Une fois la fonction exécutée, l’application Excel appelle la fonction `main` en fournissant le classeur en tant que premier paramètre. Un `ExcelScript.Workbook` doit toujours être le premier paramètre.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

Le code à l’intérieur de la fonction `main` s’exécute lors de l’exécution du script. `main` peut appeler d’autres fonctions dans le script, mais le code qui n’est pas inclus dans une fonction ne s’exécutera pas. Les scripts ne peuvent pas invoquer ou appeler d'autres scripts Office.

[Power Automate](https://flow.microsoft.com) permet de connecter des scripts dans des flux. Les données sont transmises entre les scripts et le flux entre les paramètres et les retours de la méthode `main`. L'intégration des scripts Office avec Power Automate est couverte en détail dans [Exécuter des scripts Office avec Power Automate](power-automate-integration.md).

## <a name="object-model-overview"></a>Vue d'ensemble du modèle objet

Pour écrire un script, vous devez comprendre la manière dont les API des scripts Office s’adaptent. Les composants d’un classeur sont dépendants les uns des autres. Dans de nombreux cas, ces relations correspondent à celles de l’interface utilisateur d’Excel.

- Un **classeur** contient une ou plusieurs **feuilles de calcul**.
- Une **feuille de calcul** donne accès à des cellules via **plage** objets.
- Une **plage** représente un groupe de cellules contiguës.
- Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.
- Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.
- Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.

## <a name="workbook"></a>Classeur

Chaque script est fourni avec un `workbook` objet de type `Workbook` par la fonction `main`. Il s’agit de l’objet de niveau supérieur par lequel votre script interagit avec le classeur Excel.

Le script suivant permet d’obtenir le nom de la feuille de calcul active du classeur.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a>Plages

Une plage est un groupe de cellules contiguës dans le classeur. Les scripts utilisent généralement la notation de style A1 (par exemple, **B3** pour la cellule unique dans la colonne **B** et ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et les lignes **2** via **4**) pour définir des plages.

Les plages ont trois propriétés principales : valeurs, formules et format. Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules. Ils sont accessibles via `getValues`, `getFormulas`et `getFormat`. Les valeurs et les formules peuvent être modifiées avec `setValues` et `setFormulas`, tandis que le format est un objet `RangeFormat` composé de plusieurs objets de plus petite taille définis individuellement.

Les plages utilisent des tableaux à deux dimensions pour gérer les informations. Pour plus d’informations sur la gestion des tableaux dans l’infrastructure scripts Office, consultez [Utilisation des plages](javascript-objects.md#work-with-ranges).

### <a name="range-sample"></a>Exemple de plage

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
        ["Chocolate", 10, 9.54],
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

L’exécution de ce script crée les données suivantes dans la feuille de calcul active :

:::image type="content" source="../images/range-sample.png" alt-text="Feuille de calcul contenant un enregistrement des ventes composé de lignes de valeurs, d’une colonne de formule et d’en-têtes formatés.":::

### <a name="the-types-of-range-values"></a>Les types de valeurs de plage

Chaque cellule possède une valeur. Cette valeur est la valeur sous-jacente entrée dans la cellule, qui peut être différente du texte affiché dans Excel. Par exemple, la cellule affiche la valeur « 02/05/2021 » sous forme de date, mais la valeur réelle est 44318. Cet affichage peut être modifié avec le format Nombre, mais la valeur et le type réels de la cellule ne changent que lorsqu’une nouvelle valeur est définie.

Lorsque vous utilisez la valeur de cellule, il est important d’indiquer à TypeScript la valeur attendue pour la cellule ou la plage. Une cellule contient l’un des types suivants : `string`, `number` ou `boolean`. Pour que votre script traite les valeurs renvoyées comme l’un de ces types, vous devez déclarer le type.

Le script suivant obtient le prix moyen à partir du tableau de l’exemple précédent. Notez le code `priceRange.getValues() as number[][]`. Cela [affirme](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) que le type de valeurs de plage est un `number[][]`. Toutes les valeurs de ce tableau peuvent ensuite être traitées comme des nombres dans le script.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the active worksheet.
  let sheet = workbook.getActiveWorksheet();

  // Get the "Unit Price" column. 
  // The result of calling getValues is declared to be a number[][] so that we can perform arithmetic operations.
  let priceRange = sheet.getRange("D3:D5");
  let prices = priceRange.getValues() as number[][];

  // Get the average price.
  let totalPrices = 0;
  prices.forEach((price) => totalPrices += price[0]);
  let averagePrice = totalPrices / prices.length;
  console.log(averagePrice);
}
```

## <a name="charts-tables-and-other-data-objects"></a>Graphiques, tableaux et autres objets de données

Les scripts peuvent créer et manipuler les structures de données et les visualisations dans Excel. Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore. Celles-ci sont stockées dans des collections, qui seront décrites plus loin dans cet article.

### <a name="create-a-table"></a>Créer un tableau

Créez des tableaux à l’aide de plages remplies de données. Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.

L’exemple de code suivant crée un tableau à l’aide des plages de l’exemple précédent.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

L’exécution de ce script sur la feuille de calcul avec les données précédentes crée le tableau suivant :

:::image type="content" source="../images/table-sample.png" alt-text="Feuille de calcul contenant un tableau créé depuis l’enregistrement des ventes précédent.":::

### <a name="create-a-chart"></a>Création d’un graphique (chart)

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

:::image type="content" source="../images/chart-sample.png" alt-text="Histogramme montrant les quantités pour trois des articles présents dans l’enregistrement des ventes précédent.":::

## <a name="collections"></a>Collections

Lorsqu’un objet Excel possède une collection d’un ou plusieurs objets du même type, il les stocke dans un tableau. Par exemple, un objet `Workbook` contient un `Worksheet[]`. Ce tableau est accessible à la méthode `Workbook.getWorksheets()`. Les méthodes `get` au pluriel, telles que `Worksheet.getCharts()`, renvoient l'ensemble de la collection d'objets sous forme de tableau. Vous pouvez voir ce modèle dans toutes les API Scripts Office : l’objet `Worksheet` possède une méthode `getTables()` qui renvoie un `Table[]`, l’objet `Table` possède une méthode `getColumns()` qui renvoie une `TableColumn[]`, ainsi de suite.

Le tableau retourné est un tableau normal, donc toutes les opérations normales sur les tableaux sont disponibles pour votre script. Vous pouvez également accéder aux objets individuels dans la collection à l’aide de la valeur d’index de tableau. Par exemple, `workbook.getTables()[0]` renvoie la première table de la collection. Pour plus d’informations sur l’utilisation de la fonctionnalité de tableau intégrée avec l’infrastructure Scripts Office, consultez [Utilisation des collections](javascript-objects.md#work-with-collections). 

Les objets individuels sont également accessibles à partir de la collection par le biais d'une méthode `get`. Les méthodes `get` qui sont singulières, comme `Worksheet.getTable(name)`, renvoient un seul objet et nécessitent un ID ou un nom pour l'objet spécifique. Cet ID ou nom est généralement indiqué par le script ou l’interface utilisateur d’Excel.

Le script suivant extrait toutes les tables du classeur. Il vérifie ensuite que les en-têtes sont affichés, les boutons de filtre sont visibles et le style de tableau est paramétré sur « TableStyleLight1 ».

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table collection.
  let tables = workbook.getTables();

  // Set the table formatting properties for every table.
  tables.forEach(table => {
    table.setShowHeaders(true);
    table.setShowFilterButton(true);
    table.setPredefinedTableStyle("TableStyleLight1");
  })
}
```

## <a name="add-excel-objects-with-a-script"></a>Ajouter des objets Excel à l’aide d’un script

Vous pouvez ajouter des objets document par programme, tels que des tableaux ou des graphiques, en appelant la méthode `add` correspondante disponible sur l’objet parent.

> [!IMPORTANT]
> N’ajoutez pas manuellement des objets aux tableaux de collections. Utilisez les `add` méthodes sur les objets parents par exemple, ajoutez un `Table` à une `Worksheet` avec la méthode `Worksheet.addTable`.

Le script suivant crée un tableau dans Excel sur la première feuille de calcul du classeur. Notez que la table créée est renvoyée par la méthode `addTable`.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Add a table that uses the data in A1:G10.
    let table = sheet.addTable(
      "A1:G10",
       true /* True because the table has headers. */
    );
    
    // Give the table a name for easy reference in other scripts.
    table.setName("MyTable");
}
```

> [!TIP]
> La plupart des objets Excel ont une méthode `setName` . Cela vous permet d’accéder facilement aux objets Excel plus loin dans le script ou dans d’autres scripts pour le même classeur.

### <a name="verify-an-object-exists-in-the-collection"></a>Vérifier l’existence d’un objet dans la collection

Les scripts doivent souvent vérifier si une table ou un objet similaire existe avant de continuer. Utilisez les noms donnés par les scripts ou par l'interface utilisateur d'Excel pour identifier les objets nécessaires et agir en conséquence. Les méthodes `get` renvoient `undefined` lorsque l'objet demandé ne se trouve pas dans la collection.

Le script suivant demande un tableau nommé « MyTable » et utilise une instruction `if...else` pour vérifier si le tableau a été trouvé.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable".
  let myTable = workbook.getTable("MyTable");

  // If the table is in the workbook, myTable will have a value.
  // Otherwise, the variable will be undefined and go to the else clause.
  if (myTable) {
    let worksheetName = myTable.getWorksheet().getName();
    console.log(`MyTable is on the ${worksheetName} worksheet`);
  } else {
    console.log(`MyTable is not in the workbook.`);
  }
}
```

Un modèle courant dans les scripts Office consiste à recréer un tableau, un graphique ou un autre objet à chaque exécution du script. Si vous n'avez pas besoin des anciennes données, il est préférable de supprimer l'ancien objet avant de créer le nouveau. Cela permet d’éviter les conflits de noms ou d’autres différences qui ont été introduites par d’autres utilisateurs.

Le script suivant supprime le tableau nommé « MyTable », s'il est présent, puis ajoute un nouveau tableau avec le même nom.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "MyTable" from the first worksheet.
  let sheet = workbook.getWorksheets()[0];
  let tableName = "MyTable";
  let oldTable = sheet.getTable(tableName);

  // If the table exists, remove it.
  if (oldTable) {
    oldTable.delete();
  }

  // Add a new table with the same name.
  let newTable = sheet.addTable("A1:G10", true);
  newTable.setName(tableName);
}
```

## <a name="remove-excel-objects-with-a-script"></a>Supprimer des objets Excel avec un script

Pour supprimer un objet, appelez la méthode de `delete` l’objet.

> [!NOTE]
> Comme pour l’ajout d’objets, ne supprimez pas manuellement les objets des tableaux de collections. Utilisez les méthodes `delete` sur les objets de type collection. Par exemple, supprimez un `Table` d’un `Worksheet` à l’aide d' `Table.delete`.

Le script suivant supprime la première feuille de calcul du classeur.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a>Lectures complémentaires sur le modèle d’objet

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
- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
- [Utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md)
- [Meilleures pratiques en matière de scripts Office](best-practices.md)
- [Centre de développement de scripts Office](https://developer.microsoft.com/office-scripts)
