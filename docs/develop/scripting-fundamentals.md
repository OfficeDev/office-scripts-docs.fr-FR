---
title: Principes de base des scripts pour Office Scripts dans Excel sur le web
description: Informations sur le modèle d’objet et autres concepts de base pour vous familiariser avec les scripts Office.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978718"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="45488-103">Principes de base des scripts pour Office Scripts dans Excel sur le web (préversion)</span><span class="sxs-lookup"><span data-stu-id="45488-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="45488-104">Cet article vous présente les aspects techniques de Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="45488-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="45488-105">Vous découvrirez comment les objets Excel fonctionnent ensemble et comment l’éditeur de code se synchronise avec un classeur.</span><span class="sxs-lookup"><span data-stu-id="45488-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="45488-106">Modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="45488-106">Object model</span></span>

<span data-ttu-id="45488-107">Pour comprendre les API Excel, vous devez connaître la manière dont les composants d’un classeur sont liés les uns aux autres.</span><span class="sxs-lookup"><span data-stu-id="45488-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="45488-108">Un **classeur** contient une ou plusieurs **feuilles de calcul**.</span><span class="sxs-lookup"><span data-stu-id="45488-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="45488-109">Une **feuille de calcul** donne accès à des cellules via **plage** objets.</span><span class="sxs-lookup"><span data-stu-id="45488-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="45488-110">Une **plage** représente un groupe de cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="45488-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="45488-111">Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.</span><span class="sxs-lookup"><span data-stu-id="45488-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="45488-112">Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.</span><span class="sxs-lookup"><span data-stu-id="45488-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="45488-113">Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.</span><span class="sxs-lookup"><span data-stu-id="45488-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="45488-114">Plages</span><span class="sxs-lookup"><span data-stu-id="45488-114">Ranges</span></span>

<span data-ttu-id="45488-115">Une plage est un groupe de cellules contiguës dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="45488-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="45488-116">Les scripts utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la ligne **B** et de la colonne **3** ou **C2:F4** pour les cellules des lignes **C** à **F** et des colonnes **2** à **4**) pour définir les plages.</span><span class="sxs-lookup"><span data-stu-id="45488-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="45488-117">Les plages comportent trois propriétés principales : `values`, `formulas`et `format`.</span><span class="sxs-lookup"><span data-stu-id="45488-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="45488-118">Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules.</span><span class="sxs-lookup"><span data-stu-id="45488-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="45488-119">Exemple de plage</span><span class="sxs-lookup"><span data-stu-id="45488-119">Range sample</span></span>

<span data-ttu-id="45488-120">L’exemple de code suivant montre comment créer des registres des ventes.</span><span class="sxs-lookup"><span data-stu-id="45488-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="45488-121">Le script utilise les objets `Range` pour déterminer les valeurs, les formules et les formats.</span><span class="sxs-lookup"><span data-stu-id="45488-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="45488-122">L’exécution de ce script crée les données suivantes dans la feuille de calcul active :</span><span class="sxs-lookup"><span data-stu-id="45488-122">Running this script creates the following data in the current worksheet:</span></span>

![Un registre des ventes affiche des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="45488-124">Graphiques, tableaux et autres objets de données</span><span class="sxs-lookup"><span data-stu-id="45488-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="45488-125">Les scripts peuvent créer et manipuler les structures de données et les visualisations dans Excel.</span><span class="sxs-lookup"><span data-stu-id="45488-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="45488-126">Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="45488-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="45488-127">Création d’un tableau</span><span class="sxs-lookup"><span data-stu-id="45488-127">Creating a table</span></span>

<span data-ttu-id="45488-128">Créez des tableaux à l’aide des plages de données remplies.</span><span class="sxs-lookup"><span data-stu-id="45488-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="45488-129">Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.</span><span class="sxs-lookup"><span data-stu-id="45488-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="45488-130">L’exemple de code suivant crée un tableau à l’aide des plages de l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="45488-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="45488-131">L’exécution de ce script sur la feuille de calcul avec les données précédentes crée le tableau suivant :</span><span class="sxs-lookup"><span data-stu-id="45488-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Un tableau créée à partir du registre des ventes précédent.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="45488-133">Création d’un graphique</span><span class="sxs-lookup"><span data-stu-id="45488-133">Creating a chart</span></span>

<span data-ttu-id="45488-134">Vous pouvez créer un graphique pour visualiser les données d’une plage.</span><span class="sxs-lookup"><span data-stu-id="45488-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="45488-135">Les scripts permettent des dizaines de variétés de graphiques, chacune pouvant être personnalisée pour répondre à vos besoins.</span><span class="sxs-lookup"><span data-stu-id="45488-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="45488-136">Le script suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="45488-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="45488-137">L’exécution de ce script sur la feuille de calcul avec le tableau précédent crée le graphique suivant :</span><span class="sxs-lookup"><span data-stu-id="45488-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Un histogramme montrant les quantités pour trois des articles présents dans le registre des ventes précédent.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="45488-139">Lectures complémentaires sur le modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="45488-139">Further reading on the object model</span></span>

<span data-ttu-id="45488-140">La [Documentation de référence de l’API Office Scripts](/javascript/api/office-scripts/overview) est une liste complète des objets utilisés dans Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="45488-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="45488-141">Si vous souhaitez en savoir plus, vous pouvez accéder aux informations sur la classe de votre choix en utilisant la table des matières.</span><span class="sxs-lookup"><span data-stu-id="45488-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="45488-142">Voici quelques pages fréquemment consultées.</span><span class="sxs-lookup"><span data-stu-id="45488-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="45488-143">Graphique</span><span class="sxs-lookup"><span data-stu-id="45488-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="45488-144">Commentaire</span><span class="sxs-lookup"><span data-stu-id="45488-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="45488-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="45488-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="45488-146">Range</span><span class="sxs-lookup"><span data-stu-id="45488-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="45488-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="45488-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="45488-148">Forme</span><span class="sxs-lookup"><span data-stu-id="45488-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="45488-149">Tableau</span><span class="sxs-lookup"><span data-stu-id="45488-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="45488-150">Classeur</span><span class="sxs-lookup"><span data-stu-id="45488-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="45488-151">Feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="45488-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="45488-152">Fonction `main` :</span><span class="sxs-lookup"><span data-stu-id="45488-152">`main` function</span></span>

<span data-ttu-id="45488-153">Chaque script Office doit contenir une fonction `main` avec la signature suivante, qui inclut la définition de type `Excel.RequestContext` :</span><span class="sxs-lookup"><span data-stu-id="45488-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="45488-154">Le code à l’intérieur de la fonction `main` s’exécute lors de l’exécution du script.</span><span class="sxs-lookup"><span data-stu-id="45488-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="45488-155">`main` peut appeler d’autres fonctions dans le script, mais le code qui n’est pas inclus dans une fonction ne s’exécutera pas.</span><span class="sxs-lookup"><span data-stu-id="45488-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="45488-156">Contexte</span><span class="sxs-lookup"><span data-stu-id="45488-156">Context</span></span>

<span data-ttu-id="45488-157">La fonction `main` accepte un paramètre `Excel.RequestContext`, nommé `context`.</span><span class="sxs-lookup"><span data-stu-id="45488-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="45488-158">Vous devez imaginer le `context` comme un pont entre le script et le classeur.</span><span class="sxs-lookup"><span data-stu-id="45488-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="45488-159">Le script accède au classeur avec l’objet `context` et utilise ce `context` pour envoyer et recevoir des données.</span><span class="sxs-lookup"><span data-stu-id="45488-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="45488-160">L’objet `context` est nécessaire car le script et Excel sont exécutés dans différents processus et emplacements.</span><span class="sxs-lookup"><span data-stu-id="45488-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="45488-161">Le script doit apporter des modifications ou rechercher les données du classeur dans le cloud.</span><span class="sxs-lookup"><span data-stu-id="45488-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="45488-162">L’objet `context` gère ces opérations.</span><span class="sxs-lookup"><span data-stu-id="45488-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="45488-163">Synchronisation et chargement</span><span class="sxs-lookup"><span data-stu-id="45488-163">Sync and Load</span></span>

<span data-ttu-id="45488-164">Comme le script et le classeur s’exécutent dans des emplacements différents, le transfert de données entre les deux prend du temps.</span><span class="sxs-lookup"><span data-stu-id="45488-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="45488-165">Pour améliorer les performances du script, les commandes sont mises en file d’attente jusqu’à ce que le script appelle explicitement l’opération `sync` pour synchroniser le script et le classeur.</span><span class="sxs-lookup"><span data-stu-id="45488-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="45488-166">Le script peut fonctionner de façon indépendante jusqu’à ce qu’il doive effectuer l’une des opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="45488-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="45488-167">Lire les données du classeur (en suivant une opération `load`).</span><span class="sxs-lookup"><span data-stu-id="45488-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="45488-168">Écrire les données dans le classeur (généralement quand le script est terminé).</span><span class="sxs-lookup"><span data-stu-id="45488-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="45488-169">L’image suivante montre un exemple de flux de contrôle entre le script et le classeur :</span><span class="sxs-lookup"><span data-stu-id="45488-169">The following image shows an example control flow between the script and workbook:</span></span>

![Un diagramme montrant les opérations de lecture et d’écriture effectuées dans le classeur à partir du script.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="45488-171">Synchronisation</span><span class="sxs-lookup"><span data-stu-id="45488-171">Sync</span></span>

<span data-ttu-id="45488-172">Lorsque le script a besoin de lire ou d’écrire des données dans le classeur, appelez la méthode `RequestContext.sync` comme illustré ici :</span><span class="sxs-lookup"><span data-stu-id="45488-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="45488-173">`context.sync()` est appelé implicitement à la fin d’un script.</span><span class="sxs-lookup"><span data-stu-id="45488-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="45488-174">Une fois l’opération `sync` terminée, le classeur se met à jour pour illustrer les opérations d’écriture que le script a spécifiées.</span><span class="sxs-lookup"><span data-stu-id="45488-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="45488-175">Une opération d’écriture définit une propriété sur un objet Excel (par exemple : `range.format.fill.color = "red"`) ou appelle une méthode qui modifie une propriété (par exemple : `range.format.autoFitColumns()`).</span><span class="sxs-lookup"><span data-stu-id="45488-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="45488-176">L’opération `sync` lit également les valeurs du classeur demandées par le script à l’aide d’une opération `load` (comme indiqué dans la section suivante).</span><span class="sxs-lookup"><span data-stu-id="45488-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="45488-177">La synchronisation du script avec le classeur peut prendre du temps, en fonction de votre réseau.</span><span class="sxs-lookup"><span data-stu-id="45488-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="45488-178">Vous devez diminuer le nombre d’appels `sync` pour faciliter l’exécution du script.</span><span class="sxs-lookup"><span data-stu-id="45488-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="45488-179">Charger</span><span class="sxs-lookup"><span data-stu-id="45488-179">Load</span></span>

<span data-ttu-id="45488-180">Un script doit charger les données du classeur avant de les lire.</span><span class="sxs-lookup"><span data-stu-id="45488-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="45488-181">Toutefois, le chargement fréquent de données à partir de l’intégralité du classeur réduirait considérablement la vitesse du script.</span><span class="sxs-lookup"><span data-stu-id="45488-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="45488-182">La méthode `load`, qui permet au script d’indiquer spécifiquement les données du classeur à récupérer, est plus appropriée.</span><span class="sxs-lookup"><span data-stu-id="45488-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="45488-183">La méthode `load` est disponible sur tous les objets Excel.</span><span class="sxs-lookup"><span data-stu-id="45488-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="45488-184">Le script doit charger les propriétés d’un objet avant de pouvoir les lire.</span><span class="sxs-lookup"><span data-stu-id="45488-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="45488-185">Sinon, cela entraînera une erreur.</span><span class="sxs-lookup"><span data-stu-id="45488-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="45488-186">Les exemples suivants utilisent un objet `Range` pour illustrer les trois méthodes utilisées par `load` pour charger les données.</span><span class="sxs-lookup"><span data-stu-id="45488-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="45488-187">Objectif</span><span class="sxs-lookup"><span data-stu-id="45488-187">Intent</span></span> |<span data-ttu-id="45488-188">Exemple de commande</span><span class="sxs-lookup"><span data-stu-id="45488-188">Example Command</span></span> | <span data-ttu-id="45488-189">Effet</span><span class="sxs-lookup"><span data-stu-id="45488-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="45488-190">Charger une propriété</span><span class="sxs-lookup"><span data-stu-id="45488-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="45488-191">Charge une seule propriété. Dans ce cas, le tableau à deux dimensions des valeurs dans cette plage.</span><span class="sxs-lookup"><span data-stu-id="45488-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="45488-192">Charger plusieurs propriétés</span><span class="sxs-lookup"><span data-stu-id="45488-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="45488-193">Charge toutes les propriétés d’une liste, qui sont délimitées par des virgules. Dans cet exemple, les valeurs, le nombre de lignes et le nombre de colonnes.</span><span class="sxs-lookup"><span data-stu-id="45488-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="45488-194">Tout charger</span><span class="sxs-lookup"><span data-stu-id="45488-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="45488-195">Charge toutes les propriétés de la plage.</span><span class="sxs-lookup"><span data-stu-id="45488-195">Loads all the properties on the range.</span></span> <span data-ttu-id="45488-196">Ceci n’est pas une solution recommandée, car elle ralentit le script, qui charge des données superflues.</span><span class="sxs-lookup"><span data-stu-id="45488-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="45488-197">Vous devez utiliser cette opération uniquement lorsque vous testez le script ou si vous avez besoin de toutes les propriétés de l’objet.</span><span class="sxs-lookup"><span data-stu-id="45488-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="45488-198">Le script doit appeler `context.sync()` avant de lire les valeurs chargées.</span><span class="sxs-lookup"><span data-stu-id="45488-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="45488-199">Vous pouvez également charger des propriétés sur l’ensemble d’une collection.</span><span class="sxs-lookup"><span data-stu-id="45488-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="45488-200">Chaque objet d’une collection possède une propriété `items`, qui est un tableau contenant les objets dans cette collection.</span><span class="sxs-lookup"><span data-stu-id="45488-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="45488-201">L’utilisation de `items` comme point de départ d’un appel hiérarchique (`items\myProperty`) pour que `load` charge les propriétés spécifiées sur chacun de ces éléments.</span><span class="sxs-lookup"><span data-stu-id="45488-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="45488-202">L’exemple suivant charge la propriété `resolved` sur tous les objets `Comment` dans l’objet `CommentCollection` d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="45488-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="45488-203">Si vous souhaitez en savoir plus sur l’utilisation des collections dans les scripts Office, consultez l’article [Section du tableau sur l'utilisation d'objets JavaScript intégrés dans Office Scripts](javascript-objects.md#array).</span><span class="sxs-lookup"><span data-stu-id="45488-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="45488-204">Voir également</span><span class="sxs-lookup"><span data-stu-id="45488-204">See also</span></span>

- [<span data-ttu-id="45488-205">Enregistrer, modifier et créer des scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="45488-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="45488-206">Lire les données d’un classeur avec les scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="45488-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="45488-207">Référence de l'API Office Scripts</span><span class="sxs-lookup"><span data-stu-id="45488-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="45488-208">Utilisation d’objets JavaScript intégrés dans les scripts Office</span><span class="sxs-lookup"><span data-stu-id="45488-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
