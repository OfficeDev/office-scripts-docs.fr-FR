---
title: Principes de base des scripts pour Office Scripts dans Excel sur le web
description: Informations sur le modèle d’objet et autres concepts de base pour vous familiariser avec les scripts Office.
ms.date: 05/24/2021
localization_priority: Priority
ms.openlocfilehash: 629e816ea988d6b8ffe5264c701e3a1eba6c6feb
ms.sourcegitcommit: 90ca8cdf30f2065f63938f6bb6780d024c128467
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/25/2021
ms.locfileid: "52639893"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="9ea04-103">Principes de base des scripts pour Scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="9ea04-103">Scripting fundamentals for Office Scripts in Excel on the web</span></span>

<span data-ttu-id="9ea04-104">Cet article vous présente les aspects techniques de Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="9ea04-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="9ea04-105">Vous découvrirez comment les objets Excel fonctionnent ensemble et comment l’éditeur de code se synchronise avec un classeur.</span><span class="sxs-lookup"><span data-stu-id="9ea04-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

## <a name="typescript-the-language-of-office-scripts"></a><span data-ttu-id="9ea04-106">TypeScript : langue des scripts Office</span><span class="sxs-lookup"><span data-stu-id="9ea04-106">TypeScript: The language of Office Scripts</span></span>

<span data-ttu-id="9ea04-107">Les scripts Office sont écrits dans [TypeScript](https://www.typescriptlang.org/docs/home.html), qui est un ensemble de scripts [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span><span class="sxs-lookup"><span data-stu-id="9ea04-107">Office Scripts are written in [TypeScript](https://www.typescriptlang.org/docs/home.html), which is a superset of [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript).</span></span> <span data-ttu-id="9ea04-108">Si vous avez l’expérience JavaScript, vos connaissances seront transférées, car la plupart du code est identique dans les deux langages.</span><span class="sxs-lookup"><span data-stu-id="9ea04-108">If you're familiar with JavaScript, your knowledge will carry over because much of the code is the same in both languages.</span></span> <span data-ttu-id="9ea04-109">Nous vous recommandons d'avoir des connaissances en programmation de niveau débutant avant de vous lancer dans le codage de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="9ea04-109">We recommend you have some beginner-level programming knowledge before starting your Office Scripts coding journey.</span></span> <span data-ttu-id="9ea04-110">Les ressources suivantes peuvent vous aider à comprendre l'aspect codage des scripts Office.</span><span class="sxs-lookup"><span data-stu-id="9ea04-110">The following resources can help you understand the coding side of Office Scripts.</span></span>

[!INCLUDE [Recommended coding resources](../includes/coding-basics-references.md)]

## <a name="main-function-the-scripts-starting-point"></a><span data-ttu-id="9ea04-111">Fonction `main` : point de départ du script</span><span class="sxs-lookup"><span data-stu-id="9ea04-111">`main` function: The script's starting point</span></span>

<span data-ttu-id="9ea04-112">Chaque script doit contenir une fonction `main` avec le type `ExcelScript.Workbook` comme premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="9ea04-112">Each script must contain a `main` function with the `ExcelScript.Workbook` type as its first parameter.</span></span> <span data-ttu-id="9ea04-113">Une fois la fonction exécutée, l’application Excel appelle la fonction `main` en fournissant le classeur en tant que premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="9ea04-113">When the function runs, the Excel application invokes the `main` function by providing the workbook as its first parameter.</span></span> <span data-ttu-id="9ea04-114">Un `ExcelScript.Workbook` doit toujours être le premier paramètre.</span><span class="sxs-lookup"><span data-stu-id="9ea04-114">An `ExcelScript.Workbook` should always be the first parameter.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Your code goes here
}
```

<span data-ttu-id="9ea04-115">Le code à l’intérieur de la fonction `main` s’exécute lors de l’exécution du script.</span><span class="sxs-lookup"><span data-stu-id="9ea04-115">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="9ea04-116">`main` peut appeler d’autres fonctions dans le script, mais le code qui n’est pas inclus dans une fonction ne s’exécutera pas.</span><span class="sxs-lookup"><span data-stu-id="9ea04-116">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span> <span data-ttu-id="9ea04-117">Les scripts ne peuvent pas invoquer ou appeler d'autres scripts Office.</span><span class="sxs-lookup"><span data-stu-id="9ea04-117">Scripts cannot invoke or call other Office Scripts.</span></span>

<span data-ttu-id="9ea04-118">[Power Automate](https://flow.microsoft.com) permet de connecter des scripts dans des flux.</span><span class="sxs-lookup"><span data-stu-id="9ea04-118">[Power Automate](https://flow.microsoft.com) allows you to connect scripts in flows.</span></span> <span data-ttu-id="9ea04-119">Les données sont transmises entre les scripts et le flux entre les paramètres et les retours de la méthode `main`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-119">Data is passed between the scripts and the flow through the parameters and returns of the`main` method.</span></span> <span data-ttu-id="9ea04-120">L'intégration des scripts Office avec Power Automate est couverte en détail dans [Exécuter des scripts Office avec Power Automate](power-automate-integration.md).</span><span class="sxs-lookup"><span data-stu-id="9ea04-120">How to integrate Office Scripts with Power Automate is covered in detail in [Run Office Scripts with Power Automate](power-automate-integration.md).</span></span>

## <a name="object-model-overview"></a><span data-ttu-id="9ea04-121">Vue d'ensemble du modèle objet</span><span class="sxs-lookup"><span data-stu-id="9ea04-121">Object model overview</span></span>

<span data-ttu-id="9ea04-122">Pour écrire un script, vous devez comprendre la manière dont les API des scripts Office s’adaptent.</span><span class="sxs-lookup"><span data-stu-id="9ea04-122">To write a script, you need to understand how the Office Scripts APIs fit together.</span></span> <span data-ttu-id="9ea04-123">Les composants d’un classeur sont dépendants les uns des autres.</span><span class="sxs-lookup"><span data-stu-id="9ea04-123">The components of a workbook have specific relations to one another.</span></span> <span data-ttu-id="9ea04-124">Dans de nombreux cas, ces relations correspondent à celles de l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="9ea04-124">In many ways, these relations match those of the Excel UI.</span></span>

- <span data-ttu-id="9ea04-125">Un **classeur** contient une ou plusieurs **feuilles de calcul**.</span><span class="sxs-lookup"><span data-stu-id="9ea04-125">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="9ea04-126">Une **feuille de calcul** donne accès à des cellules via **plage** objets.</span><span class="sxs-lookup"><span data-stu-id="9ea04-126">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="9ea04-127">Une **plage** représente un groupe de cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="9ea04-127">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="9ea04-128">Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.</span><span class="sxs-lookup"><span data-stu-id="9ea04-128">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="9ea04-129">Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.</span><span class="sxs-lookup"><span data-stu-id="9ea04-129">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="9ea04-130">Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.</span><span class="sxs-lookup"><span data-stu-id="9ea04-130">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

## <a name="workbook"></a><span data-ttu-id="9ea04-131">Classeur</span><span class="sxs-lookup"><span data-stu-id="9ea04-131">Workbook</span></span>

<span data-ttu-id="9ea04-132">Chaque script est fourni avec un `workbook` objet de type `Workbook` par la fonction `main`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-132">Every script is provided a `workbook` object of type `Workbook` by the `main` function.</span></span> <span data-ttu-id="9ea04-133">Il s’agit de l’objet de niveau supérieur par lequel votre script interagit avec le classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="9ea04-133">This represents the top level object through which your script interacts with the Excel workbook.</span></span>

<span data-ttu-id="9ea04-134">Le script suivant permet d’obtenir le nom de la feuille de calcul active du classeur.</span><span class="sxs-lookup"><span data-stu-id="9ea04-134">The following script gets the active worksheet from the workbook and logs its name.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Display the current worksheet's name.
    console.log(sheet.getName());
}
```

## <a name="ranges"></a><span data-ttu-id="9ea04-135">Plages</span><span class="sxs-lookup"><span data-stu-id="9ea04-135">Ranges</span></span>

<span data-ttu-id="9ea04-136">Une plage est un groupe de cellules contiguës dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="9ea04-136">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="9ea04-137">Les scripts utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la colonne **B** et de la ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et des lignes **2** à **4**) pour définir les plages.</span><span class="sxs-lookup"><span data-stu-id="9ea04-137">Scripts typically use A1-style notation (e.g., **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="9ea04-138">Les plages ont trois propriétés principales : valeurs, formules et format.</span><span class="sxs-lookup"><span data-stu-id="9ea04-138">Ranges have three core properties: values, formulas, and format.</span></span> <span data-ttu-id="9ea04-139">Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules.</span><span class="sxs-lookup"><span data-stu-id="9ea04-139">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span> <span data-ttu-id="9ea04-140">Ils sont accessibles via `getValues`, `getFormulas`et `getFormat`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-140">They are accessed through `getValues`, `getFormulas`, and `getFormat`.</span></span> <span data-ttu-id="9ea04-141">Les valeurs et les formules peuvent être modifiées avec `setValues` et `setFormulas`, tandis que le format est un objet `RangeFormat` composé de plusieurs objets de plus petite taille définis individuellement.</span><span class="sxs-lookup"><span data-stu-id="9ea04-141">Values and formulas can be changed with `setValues` and `setFormulas`, while the format is a `RangeFormat` object comprised of several smaller objects that are individually set.</span></span>

<span data-ttu-id="9ea04-142">Les plages utilisent des tableaux à deux dimensions pour gérer les informations.</span><span class="sxs-lookup"><span data-stu-id="9ea04-142">Ranges use two-dimensional arrays to manage information.</span></span> <span data-ttu-id="9ea04-143">Pour plus d’informations sur la gestion des tableaux dans l’infrastructure scripts Office, consultez [Utilisation des plages](javascript-objects.md#work-with-ranges).</span><span class="sxs-lookup"><span data-stu-id="9ea04-143">For more information on handling arrays in the Office Scripts framework, see [Work with ranges](javascript-objects.md#work-with-ranges).</span></span>

### <a name="range-sample"></a><span data-ttu-id="9ea04-144">Exemple de plage</span><span class="sxs-lookup"><span data-stu-id="9ea04-144">Range sample</span></span>

<span data-ttu-id="9ea04-145">L’exemple de code suivant montre comment créer des registres des ventes.</span><span class="sxs-lookup"><span data-stu-id="9ea04-145">The following sample shows how to create sales records.</span></span> <span data-ttu-id="9ea04-146">Ce script utilise `Range` objets pour déterminer les valeurs, les formules et les parties de la mise en forme.</span><span class="sxs-lookup"><span data-stu-id="9ea04-146">This script uses `Range` objects to set the values, formulas, and parts of the format.</span></span>

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

<span data-ttu-id="9ea04-147">L’exécution de ce script crée les données suivantes dans la feuille de calcul active :</span><span class="sxs-lookup"><span data-stu-id="9ea04-147">Running this script creates the following data in the current worksheet:</span></span>

:::image type="content" source="../images/range-sample.png" alt-text="Feuille de calcul contenant un enregistrement des ventes composé de lignes de valeurs, d’une colonne de formule et d’en-têtes formatés":::

### <a name="the-types-of-range-values"></a><span data-ttu-id="9ea04-149">Les types de valeurs de plage</span><span class="sxs-lookup"><span data-stu-id="9ea04-149">The types of Range values</span></span>

<span data-ttu-id="9ea04-150">Chaque cellule possède une valeur.</span><span class="sxs-lookup"><span data-stu-id="9ea04-150">Each cell has value.</span></span> <span data-ttu-id="9ea04-151">Cette valeur est la valeur sous-jacente entrée dans la cellule, qui peut être différente du texte affiché dans Excel.</span><span class="sxs-lookup"><span data-stu-id="9ea04-151">This value is the underlying value entered into the cell, which may be different from the text displayed in Excel.</span></span> <span data-ttu-id="9ea04-152">Par exemple, la cellule affiche la valeur « 02/05/2021 » sous forme de date, mais la valeur réelle est 44318.</span><span class="sxs-lookup"><span data-stu-id="9ea04-152">For example, you might see "5/2/2021" displayed in the cell as a date, but the actual value is 44318.</span></span> <span data-ttu-id="9ea04-153">Cet affichage peut être modifié avec le format Nombre, mais la valeur et le type réels de la cellule ne changent que lorsqu’une nouvelle valeur est définie.</span><span class="sxs-lookup"><span data-stu-id="9ea04-153">This display can be changed with the number format, but the actual value and type in the cell only changes when a new value is set.</span></span>

<span data-ttu-id="9ea04-154">Lorsque vous utilisez la valeur de cellule, il est important d’indiquer à TypeScript la valeur attendue pour la cellule ou la plage.</span><span class="sxs-lookup"><span data-stu-id="9ea04-154">When you are using the cell value, it's important to tell TypeScript what value you are expecting to get from a cell or range.</span></span> <span data-ttu-id="9ea04-155">Une cellule contient l’un des types suivants : `string`, `number` ou `boolean`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-155">A cell contains one of the following types: `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="9ea04-156">Pour que votre script traite les valeurs renvoyées comme l’un de ces types, vous devez déclarer le type.</span><span class="sxs-lookup"><span data-stu-id="9ea04-156">In order for your script to treat the returned values as one of those types, you must declare the type.</span></span>

<span data-ttu-id="9ea04-157">Le script suivant obtient le prix moyen à partir du tableau de l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="9ea04-157">The following script gets the average price from the table in the previous sample.</span></span> <span data-ttu-id="9ea04-158">Notez le code `priceRange.getValues() as number[][]`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-158">Note the code `priceRange.getValues() as number[][]`.</span></span> <span data-ttu-id="9ea04-159">Cela [affirme](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) que le type de valeurs de plage est un `number[][]`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-159">This [asserts](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#type-assertions) the type of the range values to be a `number[][]`.</span></span> <span data-ttu-id="9ea04-160">Toutes les valeurs de ce tableau peuvent ensuite être traitées comme des nombres dans le script.</span><span class="sxs-lookup"><span data-stu-id="9ea04-160">All the values in that array can then be treated as numbers in the script.</span></span>

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

## <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="9ea04-161">Graphiques, tableaux et autres objets de données</span><span class="sxs-lookup"><span data-stu-id="9ea04-161">Charts, tables, and other data objects</span></span>

<span data-ttu-id="9ea04-162">Les scripts peuvent créer et manipuler les structures de données et les visualisations dans Excel.</span><span class="sxs-lookup"><span data-stu-id="9ea04-162">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="9ea04-163">Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="9ea04-163">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span> <span data-ttu-id="9ea04-164">Celles-ci sont stockées dans des collections, qui seront décrites plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="9ea04-164">These are stored in collections, which will be discussed later in this article.</span></span>

### <a name="create-a-table"></a><span data-ttu-id="9ea04-165">Créer un tableau</span><span class="sxs-lookup"><span data-stu-id="9ea04-165">Create a table</span></span>

<span data-ttu-id="9ea04-p116">Créez des tableaux à l’aide de plages remplies de données. Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.</span><span class="sxs-lookup"><span data-stu-id="9ea04-p116">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="9ea04-168">L’exemple de code suivant crée un tableau à l’aide des plages de l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="9ea04-168">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Add a table that has headers using the data from B2:E5.
    sheet.addTable("B2:E5", true);
}
```

<span data-ttu-id="9ea04-169">L’exécution de ce script sur la feuille de calcul avec les données précédentes crée le tableau suivant :</span><span class="sxs-lookup"><span data-stu-id="9ea04-169">Running this script on the worksheet with the previous data creates the following table:</span></span>

:::image type="content" source="../images/table-sample.png" alt-text="Feuille de calcul contenant un tableau créé depuis l’enregistrement des ventes précédent":::

### <a name="create-a-chart"></a><span data-ttu-id="9ea04-171">Création d’un graphique (chart)</span><span class="sxs-lookup"><span data-stu-id="9ea04-171">Create a chart</span></span>

<span data-ttu-id="9ea04-172">Vous pouvez créer un graphique pour visualiser les données d’une plage.</span><span class="sxs-lookup"><span data-stu-id="9ea04-172">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="9ea04-173">Les scripts permettent des dizaines de variétés de graphiques, chacune pouvant être personnalisée pour répondre à vos besoins.</span><span class="sxs-lookup"><span data-stu-id="9ea04-173">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="9ea04-174">Le script suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="9ea04-174">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

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

<span data-ttu-id="9ea04-175">L’exécution de ce script sur la feuille de calcul avec le tableau précédent crée le graphique suivant :</span><span class="sxs-lookup"><span data-stu-id="9ea04-175">Running this script on the worksheet with the previous table creates the following chart:</span></span>

:::image type="content" source="../images/chart-sample.png" alt-text="Histogramme montrant les quantités pour trois des articles présents dans l’enregistrement des ventes précédent":::

## <a name="collections"></a><span data-ttu-id="9ea04-177">Collections</span><span class="sxs-lookup"><span data-stu-id="9ea04-177">Collections</span></span>

<span data-ttu-id="9ea04-178">Lorsqu’un objet Excel possède une collection d’un ou plusieurs objets du même type, il les stocke dans un tableau.</span><span class="sxs-lookup"><span data-stu-id="9ea04-178">When an Excel object has a collection of one or more objects of the same type, it stores them in an array.</span></span> <span data-ttu-id="9ea04-179">Par exemple, un objet `Workbook` contient un `Worksheet[]`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-179">For example, a `Workbook` object contains a `Worksheet[]`.</span></span> <span data-ttu-id="9ea04-180">Ce tableau est accessible à la méthode `Workbook.getWorksheets()`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-180">This array is accessed by the `Workbook.getWorksheets()` method.</span></span> <span data-ttu-id="9ea04-181">Les méthodes `get` au pluriel, telles que `Worksheet.getCharts()`, renvoient l'ensemble de la collection d'objets sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="9ea04-181">`get` methods that are plural, such as `Worksheet.getCharts()`, return the entire object collection as an array.</span></span> <span data-ttu-id="9ea04-182">Vous pouvez voir ce modèle dans toutes les API Scripts Office : l’objet `Worksheet` possède une méthode `getTables()` qui renvoie un `Table[]`, l’objet `Table` possède une méthode `getColumns()` qui renvoie une `TableColumn[]`, ainsi de suite.</span><span class="sxs-lookup"><span data-stu-id="9ea04-182">You'll see this pattern throughout the Office Scripts APIs: the `Worksheet` object has a `getTables()` method that returns a `Table[]`, the `Table` object has a `getColumns()` method that returns a `TableColumn[]`, as so on.</span></span>

<span data-ttu-id="9ea04-183">Le tableau retourné est un tableau normal, donc toutes les opérations normales sur les tableaux sont disponibles pour votre script.</span><span class="sxs-lookup"><span data-stu-id="9ea04-183">The returned array is a normal array, so all the regular array operations are available for your script.</span></span> <span data-ttu-id="9ea04-184">Vous pouvez également accéder aux objets individuels dans la collection à l’aide de la valeur d’index de tableau.</span><span class="sxs-lookup"><span data-stu-id="9ea04-184">You can also access individual objects within the collection using the array index value.</span></span> <span data-ttu-id="9ea04-185">Par exemple, `workbook.getTables()[0]` renvoie la première table de la collection.</span><span class="sxs-lookup"><span data-stu-id="9ea04-185">For example, `workbook.getTables()[0]` returns the first table in the collection.</span></span> <span data-ttu-id="9ea04-186">Pour plus d’informations sur l’utilisation de la fonctionnalité de tableau intégrée avec l’infrastructure Scripts Office, consultez [Utilisation des collections](javascript-objects.md#work-with-collections).</span><span class="sxs-lookup"><span data-stu-id="9ea04-186">For more information on using the built-in array functionality with the Office Scripts framework, see [Work with collections](javascript-objects.md#work-with-collections).</span></span> 

<span data-ttu-id="9ea04-187">Les objets individuels sont également accessibles à partir de la collection par le biais d'une méthode `get`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-187">Individual objects are also accessed from the collection through a `get` method.</span></span> <span data-ttu-id="9ea04-188">Les méthodes `get` qui sont singulières, comme `Worksheet.getTable(name)`, renvoient un seul objet et nécessitent un ID ou un nom pour l'objet spécifique.</span><span class="sxs-lookup"><span data-stu-id="9ea04-188">`get` methods that are singular, such as `Worksheet.getTable(name)`, return a single object and require an ID or name for the specific object.</span></span> <span data-ttu-id="9ea04-189">Cet ID ou nom est généralement indiqué par le script ou l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="9ea04-189">This ID or name is usually set by the script or through the Excel UI.</span></span>

<span data-ttu-id="9ea04-p121">Le script suivant extrait toutes les tables du classeur. Il vérifie ensuite que les en-têtes sont affichés, les boutons de filtre sont visibles et le style de tableau est paramétré sur « TableStyleLight1 ».</span><span class="sxs-lookup"><span data-stu-id="9ea04-p121">The following script gets all tables in the workbook. It then ensures the headers are displays, the filter buttons are visible, and the table style is set to "TableStyleLight1".</span></span>

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

## <a name="add-excel-objects-with-a-script"></a><span data-ttu-id="9ea04-192">Ajouter des objets Excel à l’aide d’un script</span><span class="sxs-lookup"><span data-stu-id="9ea04-192">Add Excel objects with a script</span></span>

<span data-ttu-id="9ea04-193">Vous pouvez ajouter des objets document par programme, tels que des tableaux ou des graphiques, en appelant la méthode `add` correspondante disponible sur l’objet parent.</span><span class="sxs-lookup"><span data-stu-id="9ea04-193">You can programmatically add document objects, such as tables or charts, by calling the corresponding `add` method available on the parent object.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9ea04-194">N’ajoutez pas manuellement des objets aux tableaux de collections.</span><span class="sxs-lookup"><span data-stu-id="9ea04-194">Do not manually add objects to collection arrays.</span></span> <span data-ttu-id="9ea04-195">Utilisez les `add` méthodes sur les objets parents par exemple, ajoutez un `Table` à une `Worksheet` avec la méthode `Worksheet.addTable`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-195">Use the `add` methods on the parent objects For example, add a `Table` to a `Worksheet` with the `Worksheet.addTable` method.</span></span>

<span data-ttu-id="9ea04-196">Le script suivant crée un tableau dans Excel sur la première feuille de calcul du classeur.</span><span class="sxs-lookup"><span data-stu-id="9ea04-196">The following script creates a table in Excel on the first worksheet in the workbook.</span></span> <span data-ttu-id="9ea04-197">Notez que la table créée est renvoyée par la méthode `addTable`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-197">Note that the created table is returned by the `addTable` method.</span></span>

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
> <span data-ttu-id="9ea04-198">La plupart des objets Excel ont une méthode `setName`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-198">Most Excel objects have a `setName` method.</span></span> <span data-ttu-id="9ea04-199">Cela vous donne un moyen facile d'accéder aux objets Excel plus tard dans le script ou dans d'autres scripts pour le même classeur.</span><span class="sxs-lookup"><span data-stu-id="9ea04-199">This gives you an easy way to access Excel objects later in the script or in other scripts for the same workbook.</span></span>

### <a name="verify-an-object-exists-in-the-collection"></a><span data-ttu-id="9ea04-200">Vérifier l’existence d’un objet dans la collection</span><span class="sxs-lookup"><span data-stu-id="9ea04-200">Verify an object exists in the collection</span></span>

<span data-ttu-id="9ea04-201">Les scripts doivent souvent vérifier si une table ou un objet similaire existe avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="9ea04-201">Scripts often need to check if a table or similar object exists before continuing.</span></span> <span data-ttu-id="9ea04-202">Utilisez les noms donnés par les scripts ou par l'interface utilisateur d'Excel pour identifier les objets nécessaires et agir en conséquence.</span><span class="sxs-lookup"><span data-stu-id="9ea04-202">Use the names given by scripts or through the Excel UI to identify necessary objects and act accordingly.</span></span> <span data-ttu-id="9ea04-203">Les méthodes `get` renvoient `undefined` lorsque l'objet demandé ne se trouve pas dans la collection.</span><span class="sxs-lookup"><span data-stu-id="9ea04-203">`get` methods return `undefined` when the requested object is not in the collection.</span></span>

<span data-ttu-id="9ea04-204">Le script suivant demande un tableau nommé « MyTable » et utilise une instruction `if...else` pour vérifier si le tableau a été trouvé.</span><span class="sxs-lookup"><span data-stu-id="9ea04-204">The following script requests a table named "MyTable" and uses an `if...else` statement to check if the table was found.</span></span>

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

<span data-ttu-id="9ea04-205">Un modèle courant dans les scripts Office consiste à recréer un tableau, un graphique ou un autre objet à chaque exécution du script.</span><span class="sxs-lookup"><span data-stu-id="9ea04-205">A common pattern in Office Scripts is to recreate a table, chart, or other object every time the script is run.</span></span> <span data-ttu-id="9ea04-206">Si vous n'avez pas besoin des anciennes données, il est préférable de supprimer l'ancien objet avant de créer le nouveau.</span><span class="sxs-lookup"><span data-stu-id="9ea04-206">If you don't need the old data, it's best to delete the old object before creating the new one.</span></span> <span data-ttu-id="9ea04-207">Cela permet d’éviter les conflits de noms ou d’autres différences qui ont été introduites par d’autres utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="9ea04-207">This avoids name conflicts or other differences that may have been introduced by other users.</span></span>

<span data-ttu-id="9ea04-208">Le script suivant supprime le tableau nommé « MyTable », s'il est présent, puis ajoute un nouveau tableau avec le même nom.</span><span class="sxs-lookup"><span data-stu-id="9ea04-208">The following script removes the table named "MyTable", if it is present, then adds a new table with the same name.</span></span>

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

## <a name="remove-excel-objects-with-a-script"></a><span data-ttu-id="9ea04-209">Supprimer des objets Excel avec un script</span><span class="sxs-lookup"><span data-stu-id="9ea04-209">Remove Excel objects with a script</span></span>

<span data-ttu-id="9ea04-210">Pour supprimer un objet, appelez la méthode de `delete` l’objet.</span><span class="sxs-lookup"><span data-stu-id="9ea04-210">To delete an object, call the object's `delete` method.</span></span>

> [!NOTE]
> <span data-ttu-id="9ea04-211">Comme pour l’ajout d’objets, ne supprimez pas manuellement les objets des tableaux de collections.</span><span class="sxs-lookup"><span data-stu-id="9ea04-211">As with adding objects, do not manually remove objects from collection arrays.</span></span> <span data-ttu-id="9ea04-212">Utilisez les méthodes `delete` sur les objets de type collection.</span><span class="sxs-lookup"><span data-stu-id="9ea04-212">Use the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="9ea04-213">Par exemple, supprimez un `Table` d’un `Worksheet` à l’aide d' `Table.delete`.</span><span class="sxs-lookup"><span data-stu-id="9ea04-213">For example, remove a `Table` from a `Worksheet` using `Table.delete`.</span></span>

<span data-ttu-id="9ea04-214">Le script suivant supprime la première feuille de calcul du classeur.</span><span class="sxs-lookup"><span data-stu-id="9ea04-214">The following script removes the first worksheet in the workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Get first worksheet.
    let sheet = workbook.getWorksheets()[0];

    // Remove that worksheet from the workbook.
    sheet.delete();
}
```

## <a name="further-reading-on-the-object-model"></a><span data-ttu-id="9ea04-215">Lectures complémentaires sur le modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="9ea04-215">Further reading on the object model</span></span>

<span data-ttu-id="9ea04-216">La [Documentation de référence de l’API Office Scripts](/javascript/api/office-scripts/overview) est une liste complète des objets utilisés dans Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="9ea04-216">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="9ea04-217">Si vous souhaitez en savoir plus, vous pouvez accéder aux informations sur la classe de votre choix en utilisant la table des matières.</span><span class="sxs-lookup"><span data-stu-id="9ea04-217">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="9ea04-218">Voici quelques pages fréquemment consultées.</span><span class="sxs-lookup"><span data-stu-id="9ea04-218">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="9ea04-219">Graphique</span><span class="sxs-lookup"><span data-stu-id="9ea04-219">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart)
- [<span data-ttu-id="9ea04-220">Commentaire</span><span class="sxs-lookup"><span data-stu-id="9ea04-220">Comment</span></span>](/javascript/api/office-scripts/excelscript/excelscript.comment)
- [<span data-ttu-id="9ea04-221">PivotTable</span><span class="sxs-lookup"><span data-stu-id="9ea04-221">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable)
- [<span data-ttu-id="9ea04-222">Range</span><span class="sxs-lookup"><span data-stu-id="9ea04-222">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range)
- [<span data-ttu-id="9ea04-223">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="9ea04-223">RangeFormat</span></span>](/javascript/api/office-scripts/excelscript/excelscript.rangeformat)
- [<span data-ttu-id="9ea04-224">Forme</span><span class="sxs-lookup"><span data-stu-id="9ea04-224">Shape</span></span>](/javascript/api/office-scripts/excelscript/excelscript.shape)
- [<span data-ttu-id="9ea04-225">Tableau</span><span class="sxs-lookup"><span data-stu-id="9ea04-225">Table</span></span>](/javascript/api/office-scripts/excelscript/excelscript.table)
- [<span data-ttu-id="9ea04-226">Classeur</span><span class="sxs-lookup"><span data-stu-id="9ea04-226">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook)
- [<span data-ttu-id="9ea04-227">Feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="9ea04-227">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet)

## <a name="see-also"></a><span data-ttu-id="9ea04-228">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9ea04-228">See also</span></span>

- [<span data-ttu-id="9ea04-229">Enregistrer, modifier et créer des scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="9ea04-229">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="9ea04-230">Lire les données d’un classeur avec les scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="9ea04-230">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="9ea04-231">Référence de l'API Office Scripts</span><span class="sxs-lookup"><span data-stu-id="9ea04-231">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="9ea04-232">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="9ea04-232">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
- [<span data-ttu-id="9ea04-233">Meilleures pratiques en matière de scripts Office</span><span class="sxs-lookup"><span data-stu-id="9ea04-233">Best practices in Office Scripts</span></span>](best-practices.md)
