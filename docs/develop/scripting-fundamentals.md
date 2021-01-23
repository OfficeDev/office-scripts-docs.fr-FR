---
title: Principes de base des scripts pour Office Scripts dans Excel sur le web
description: Informations sur le modèle d’objet et autres concepts de base pour vous familiariser avec les scripts Office.
ms.date: 07/08/2020
localization_priority: Priority
ms.openlocfilehash: acbeec69a5d9ae9e3ebfa95c9070033d1cca2265
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933272"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="fc045-103">Principes de base des scripts pour Office Scripts dans Excel sur le web (préversion)</span><span class="sxs-lookup"><span data-stu-id="fc045-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="fc045-104">Cet article vous présente les aspects techniques de Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="fc045-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="fc045-105">Vous découvrirez comment les objets Excel fonctionnent ensemble et comment l’éditeur de code se synchronise avec un classeur.</span><span class="sxs-lookup"><span data-stu-id="fc045-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="main-function"></a><span data-ttu-id="fc045-106">Fonction `main` :</span><span class="sxs-lookup"><span data-stu-id="fc045-106">`main` function</span></span>

<span data-ttu-id="fc045-107">Chaque script Office doit contenir une fonction `main` avec le type `ExcelScript.Workbook` comme premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="fc045-107">Each Office Script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="fc045-108">Une fois la fonction exécutée, l’application Excel appelle cette fonction `main` en fournissant le classeur en tant que premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="fc045-108">When the function is executed, the Excel application invokes this `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="fc045-109">Par conséquent, il est important de ne pas modifier la signature de base de la fonction `main` une fois que vous avez enregistré le script ou créé un nouveau script à partir de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="fc045-109">Hence, it is important to not modify the basic signature of the `main` function once you have either recorded the script or created a new script from the code editor.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="fc045-110">Le code à l’intérieur de la fonction `main` s’exécute lors de l’exécution du script.</span><span class="sxs-lookup"><span data-stu-id="fc045-110">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="fc045-111">`main` peut appeler d’autres fonctions dans le script, mais le code qui n’est pas inclus dans une fonction ne s’exécutera pas.</span><span class="sxs-lookup"><span data-stu-id="fc045-111">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

> [!CAUTION]
> <span data-ttu-id="fc045-112">Si votre fonction `main` ressemble à ceci : `async function main(context: Excel.RequestContext)`, cela veut dire que votre script utilise l’ancien modèle API asynchrone.</span><span class="sxs-lookup"><span data-stu-id="fc045-112">If your `main` function looks like `async function main(context: Excel.RequestContext)`, your script is using the older async API model.</span></span> <span data-ttu-id="fc045-113">Si vous souhaitez en savoir plus (notamment sur la conversion de votre script vers le modèle API actuel), veuillez consulter l’article [Prendre en charge les anciens scripts Office qui utilisent des API asynchrones](excel-async-model.md).</span><span class="sxs-lookup"><span data-stu-id="fc045-113">For more information (including how to convert your script to the current API model), refer to [Support older Office Scripts that use the Async APIs](excel-async-model.md).</span></span>

## <a name="object-model"></a><span data-ttu-id="fc045-114">Modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="fc045-114">Object model</span></span>

<span data-ttu-id="fc045-115">Pour écrire un script, vous devez comprendre la manière dont les API de script Office s’adaptent.</span><span class="sxs-lookup"><span data-stu-id="fc045-115">To write a script, you need to understand how the Office Script APIs fit together.</span></span> <span data-ttu-id="fc045-116">Les composants d’un classeur sont dépendants les uns des autres.</span><span class="sxs-lookup"><span data-stu-id="fc045-116">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="fc045-117">Dans de nombreux cas, ces relations correspondent à celles de l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="fc045-117">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="fc045-118">Un **classeur** contient une ou plusieurs **feuilles de calcul**.</span><span class="sxs-lookup"><span data-stu-id="fc045-118">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="fc045-119">Une **feuille de calcul** donne accès à des cellules via **plage** objets.</span><span class="sxs-lookup"><span data-stu-id="fc045-119">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="fc045-120">Une **plage** représente un groupe de cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="fc045-120">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="fc045-121">Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.</span><span class="sxs-lookup"><span data-stu-id="fc045-121">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="fc045-122">Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.</span><span class="sxs-lookup"><span data-stu-id="fc045-122">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="fc045-123">Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.</span><span class="sxs-lookup"><span data-stu-id="fc045-123">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="workbook"></a><span data-ttu-id="fc045-124">Classeur</span><span class="sxs-lookup"><span data-stu-id="fc045-124">Workbook</span></span>

<span data-ttu-id="fc045-125">Chaque script est fourni avec un `workbook` objet de type `Workbook` par la fonction `main`.</span><span class="sxs-lookup"><span data-stu-id="fc045-125">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="fc045-126">Il s’agit de l’objet de niveau supérieur par lequel votre script interagit avec le classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="fc045-126">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="fc045-127">Le script suivant permet d’obtenir le nom de la feuille de calcul active du classeur.</span><span class="sxs-lookup"><span data-stu-id="fc045-127">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

### <a name="ranges"></a><span data-ttu-id="fc045-128">Plages</span><span class="sxs-lookup"><span data-stu-id="fc045-128">Ranges</span></span>

<span data-ttu-id="fc045-129">Une plage est un groupe de cellules contiguës dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="fc045-129">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="fc045-130">Les scripts utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la colonne **B** et de la ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et des lignes **2** à **4**) pour définir les plages.</span><span class="sxs-lookup"><span data-stu-id="fc045-130">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="fc045-131">Les plages ont trois propriétés principales : valeurs, formules et format.</span><span class="sxs-lookup"><span data-stu-id="fc045-131">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="fc045-132">Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules.</span><span class="sxs-lookup"><span data-stu-id="fc045-132">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="fc045-133">Ils sont accessibles via `getValues`, `getFormulas`et `getFormat`.</span><span class="sxs-lookup"><span data-stu-id="fc045-133">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="fc045-134">Les valeurs et les formules peuvent être modifiées avec `setValues` et `setFormulas`, tandis que le format est un objet `RangeFormat` composé de plusieurs objets de plus petite taille définis individuellement.</span><span class="sxs-lookup"><span data-stu-id="fc045-134">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="fc045-135">Les plages utilisent des tableaux à deux dimensions pour gérer les informations.</span><span class="sxs-lookup"><span data-stu-id="fc045-135">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="fc045-136">Pour plus d’informations sur la gestion de ces tableaux dans la structure de scripts Office, consultez la section [utilisation des plages de la section utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md#working-with-ranges).</span><span class="sxs-lookup"><span data-stu-id="fc045-136">Read the [Working with ranges section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-ranges) for more information on handling those arrays in the Office Scripts framework.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="fc045-137">Exemple de plage</span><span class="sxs-lookup"><span data-stu-id="fc045-137">Range sample</span></span>

<span data-ttu-id="fc045-138">L’exemple de code suivant montre comment créer des registres des ventes.</span><span class="sxs-lookup"><span data-stu-id="fc045-138">The following sample shows how to create sales records.</span></span> <span data-ttu-id="fc045-139">Ce script utilise `Range` objets pour déterminer les valeurs, les formules et les parties de la mise en forme.</span><span class="sxs-lookup"><span data-stu-id="fc045-139">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="fc045-140">L’exécution de ce script crée les données suivantes dans la feuille de calcul active :</span><span class="sxs-lookup"><span data-stu-id="fc045-140">Running this script creates the following data in the current worksheet:</span></span>

![Un registre des ventes affiche des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="fc045-142">Graphiques, tableaux et autres objets de données</span><span class="sxs-lookup"><span data-stu-id="fc045-142">Charts, tables, and other data objects</span></span>

<span data-ttu-id="fc045-143">Les scripts peuvent créer et manipuler les structures de données et les visualisations dans Excel.</span><span class="sxs-lookup"><span data-stu-id="fc045-143">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="fc045-144">Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="fc045-144">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="fc045-145">Celles-ci sont stockées dans des collections, qui seront décrites plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="fc045-145">These are stored in collections, which will be discussed later in this article.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="fc045-146">Création d’un tableau</span><span class="sxs-lookup"><span data-stu-id="fc045-146">Creating a table</span></span>

<span data-ttu-id="fc045-147">Créez des tableaux à l’aide des plages de données remplies.</span><span class="sxs-lookup"><span data-stu-id="fc045-147">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="fc045-148">Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.</span><span class="sxs-lookup"><span data-stu-id="fc045-148">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="fc045-149">L’exemple de code suivant crée un tableau à l’aide des plages de l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="fc045-149">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="fc045-150">L’exécution de ce script sur la feuille de calcul avec les données précédentes crée le tableau suivant :</span><span class="sxs-lookup"><span data-stu-id="fc045-150">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Un tableau créée à partir du registre des ventes précédent.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="fc045-152">Création d’un graphique</span><span class="sxs-lookup"><span data-stu-id="fc045-152">Creating a chart</span></span>

<span data-ttu-id="fc045-153">Vous pouvez créer un graphique pour visualiser les données d’une plage.</span><span class="sxs-lookup"><span data-stu-id="fc045-153">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="fc045-154">Les scripts permettent des dizaines de variétés de graphiques, chacune pouvant être personnalisée pour répondre à vos besoins.</span><span class="sxs-lookup"><span data-stu-id="fc045-154">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="fc045-155">Le script suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="fc045-155">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="fc045-156">L’exécution de ce script sur la feuille de calcul avec le tableau précédent crée le graphique suivant :</span><span class="sxs-lookup"><span data-stu-id="fc045-156">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Un histogramme montrant les quantités pour trois des articles présents dans le registre des ventes précédent.](../images/chart-sample.png)

### <a name="collections-and-other-object-relations"></a><span data-ttu-id="fc045-158">Collections et autres relations d’objets</span><span class="sxs-lookup"><span data-stu-id="fc045-158">Collections and other object relations</span></span>

<span data-ttu-id="fc045-159">Tout objet enfant est accessible via son objet parent.</span><span class="sxs-lookup"><span data-stu-id="fc045-159">Any child object can be accessed through its parent object.</span></span> <span data-ttu-id="fc045-160">Par exemple, vous pouvez lire `Worksheets` à partir de l’objet `Workbook`.</span><span class="sxs-lookup"><span data-stu-id="fc045-160">For example, you can read `Worksheets` from the `Workbook` object.</span></span> <span data-ttu-id="fc045-161">Il y aura une méthode `get` associée sur la classe parente (par exemple, `Workbook.getWorksheets()` ou `Workbook.getWorksheet(name)`).</span><span class="sxs-lookup"><span data-stu-id="fc045-161">There will be a related `get` method on the parent class that (e.g., `Workbook.getWorksheets()` or `Workbook.getWorksheet(name)`).</span></span> <span data-ttu-id="fc045-162">`get` les méthodes qui sont singulières renvoient un objet unique et nécessitent un ID ou un nom pour l’objet spécifique (par exemple, le nom d’une feuille de calcul).</span><span class="sxs-lookup"><span data-stu-id="fc045-162">`get` methods that are singular return a single object and require an ID or name for the specific object (such as the name of a worksheet).</span></span> <span data-ttu-id="fc045-163">`get` les méthodes qui permettent de renvoyer l’ensemble de la collection d’objets sous la forme d’une matrice.</span><span class="sxs-lookup"><span data-stu-id="fc045-163">`get` methods that are plural return the entire object collection as an array.</span></span> <span data-ttu-id="fc045-164">Si la collection est vide, vous obtenez une matrice vide (`[]`).</span><span class="sxs-lookup"><span data-stu-id="fc045-164">If the collection is empty, you'll get an empty array (`[]`).</span></span>

<span data-ttu-id="fc045-165">Une fois la collection récupérée, vous pouvez utiliser des opérations de tableau régulières, telles que l’acquisition de ses `length` ou utiliser des `for`, `for..of``while` des boucles pour l’itération ou utiliser des méthodes matricielles telles que les `map``forEach`.</span><span class="sxs-lookup"><span data-stu-id="fc045-165">Once the collection is retrieved, you can use regular array operations such as getting its `length` or use `for`, `for..of`, `while` loops for iteration or use TypeScript array methods such as `map`, `forEach` on them.</span></span> <span data-ttu-id="fc045-166">Vous pouvez également accéder aux objets individuels dans la collection à l’aide de la valeur d’index de tableau.</span><span class="sxs-lookup"><span data-stu-id="fc045-166">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="fc045-167">Par exemple, `workbook.getTables()[0]` renvoie la première table de la collection.</span><span class="sxs-lookup"><span data-stu-id="fc045-167">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="fc045-168">Pour en savoir plus sur l’utilisation de la fonctionnalité de tableau intégrée avec l’infrastructure de scripts Office, consultez la section [utilisation des collections de utilisation d’objets JavaScript intégrés dans les scripts Office](javascript-objects.md#working-with-collections).</span><span class="sxs-lookup"><span data-stu-id="fc045-168">Read the [Working with collections section of Using built-in JavaScript objects in Office Scripts](javascript-objects.md#working-with-collections) to learn more about using built-in array functionality with the Office Scripts framework.</span></span>

<span data-ttu-id="fc045-169">Le script suivant extrait toutes les tables du classeur.</span><span class="sxs-lookup"><span data-stu-id="fc045-169">The following script gets all tables in the workbook.</span></span> <span data-ttu-id="fc045-170">Il vérifie ensuite que les en-têtes sont affichés, les boutons de filtre sont visibles et le style de tableau est paramétré sur « TableStyleLight1 ».</span><span class="sxs-lookup"><span data-stu-id="fc045-170">It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

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

#### <a name="adding-excel-objects-with-a-script"></a><span data-ttu-id="fc045-171">Ajout d’objets Excel à l’aide d’un script</span><span class="sxs-lookup"><span data-stu-id="fc045-171">Adding Excel objects with a script</span></span>

<span data-ttu-id="fc045-172">Vous pouvez ajouter des objets document par programme, tels que des tableaux ou des graphiques, en appelant la méthode `add` correspondante disponible sur l’objet parent.</span><span class="sxs-lookup"><span data-stu-id="fc045-172">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!NOTE]
> <span data-ttu-id="fc045-173">N’ajoutez pas manuellement des objets aux tableaux de collections.</span><span class="sxs-lookup"><span data-stu-id="fc045-173">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="fc045-174">Utilisez les `add` méthodes sur les objets parents par exemple, ajoutez un `Table` à une `Worksheet` avec la méthode `Worksheet.addTable`.</span><span class="sxs-lookup"><span data-stu-id="fc045-174">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="fc045-175">Le script suivant crée un tableau dans Excel sur la première feuille de calcul du classeur.</span><span class="sxs-lookup"><span data-stu-id="fc045-175">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="fc045-176">Notez que la table créée est renvoyée par la méthode `addTable`.</span><span class="sxs-lookup"><span data-stu-id="fc045-176">Note that the created table is returned by the `addTable` method.</span></span>

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

## <a name="removing-excel-objects-with-a-script"></a><span data-ttu-id="fc045-177">Suppression d’objets Excel à l’aide d’un script</span><span class="sxs-lookup"><span data-stu-id="fc045-177">Removing Excel objects with a script</span></span>

<span data-ttu-id="fc045-178">Pour supprimer un objet, appelez la méthode de `delete` l’objet.</span><span class="sxs-lookup"><span data-stu-id="fc045-178">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="fc045-179">Comme pour l’ajout d’objets, ne supprimez pas manuellement les objets des tableaux de collections.</span><span class="sxs-lookup"><span data-stu-id="fc045-179">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="fc045-180">Utilisez les méthodes `delete` sur les objets de type collection.</span><span class="sxs-lookup"><span data-stu-id="fc045-180">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="fc045-181">Par exemple, supprimez un `Table` d’un `Worksheet` à l’aide d' `Table.delete`.</span><span class="sxs-lookup"><span data-stu-id="fc045-181">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="fc045-182">Le script suivant supprime la première feuille de calcul du classeur.</span><span class="sxs-lookup"><span data-stu-id="fc045-182">The following script removes the first worksheet in the workbook.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="fc045-183">Lectures complémentaires sur le modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="fc045-183">Further reading on the object model</span></span>

<span data-ttu-id="fc045-184">La [Documentation de référence de l’API Office Scripts](/javascript/api/office-scripts/overview) est une liste complète des objets utilisés dans Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="fc045-184">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="fc045-185">Si vous souhaitez en savoir plus, vous pouvez accéder aux informations sur la classe de votre choix en utilisant la table des matières.</span><span class="sxs-lookup"><span data-stu-id="fc045-185">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="fc045-186">Voici quelques pages fréquemment consultées.</span><span class="sxs-lookup"><span data-stu-id="fc045-186">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="fc045-187">Graphique</span><span class="sxs-lookup"><span data-stu-id="fc045-187">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="fc045-188">Commentaire</span><span class="sxs-lookup"><span data-stu-id="fc045-188">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="fc045-189">PivotTable</span><span class="sxs-lookup"><span data-stu-id="fc045-189">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="fc045-190">Range</span><span class="sxs-lookup"><span data-stu-id="fc045-190">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="fc045-191">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="fc045-191">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="fc045-192">Forme</span><span class="sxs-lookup"><span data-stu-id="fc045-192">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="fc045-193">Tableau</span><span class="sxs-lookup"><span data-stu-id="fc045-193">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="fc045-194">Classeur</span><span class="sxs-lookup"><span data-stu-id="fc045-194">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="fc045-195">Feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="fc045-195">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="fc045-196">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fc045-196">See also</span></span>

- [<span data-ttu-id="fc045-197">Enregistrer, modifier et créer des scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="fc045-197">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="fc045-198">Lire les données d’un classeur avec les scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="fc045-198">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="fc045-199">Référence de l'API Office Scripts</span><span class="sxs-lookup"><span data-stu-id="fc045-199">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="fc045-200">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="fc045-200">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
