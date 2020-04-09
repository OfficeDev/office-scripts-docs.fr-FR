---
title: Exemples de scripts pour les scripts Office dans Excel sur le Web
description: Collection d’exemples de code à utiliser avec des scripts Office dans Excel sur le Web.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: abf6b87b63ad027cca8ee5c947b687f54815409c
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191009"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="230df-103">Exemples de scripts pour les scripts Office dans Excel sur le Web (aperçu)</span><span class="sxs-lookup"><span data-stu-id="230df-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="230df-104">Les exemples suivants sont des scripts simples que vous pouvez essayer dans vos propres classeurs.</span><span class="sxs-lookup"><span data-stu-id="230df-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="230df-105">Pour les utiliser dans Excel sur le Web :</span><span class="sxs-lookup"><span data-stu-id="230df-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="230df-106">Ouvrez l’onglet **Automatiser**.</span><span class="sxs-lookup"><span data-stu-id="230df-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="230df-107">Appuyez sur **éditeur de code**.</span><span class="sxs-lookup"><span data-stu-id="230df-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="230df-108">Appuyez sur **nouveau script** dans le volet Office de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="230df-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="230df-109">Remplacez l’intégralité du script par l’exemple de votre choix.</span><span class="sxs-lookup"><span data-stu-id="230df-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="230df-110">Appuyez sur **exécuter** dans le volet Office de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="230df-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="230df-111">Concepts de base des scripts</span><span class="sxs-lookup"><span data-stu-id="230df-111">Scripting basics</span></span>

<span data-ttu-id="230df-112">Ces exemples illustrent des blocs de construction fondamentaux pour les scripts Office.</span><span class="sxs-lookup"><span data-stu-id="230df-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="230df-113">Ajoutez-les à vos scripts pour étendre votre solution et résoudre les problèmes courants.</span><span class="sxs-lookup"><span data-stu-id="230df-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="230df-114">Lecture et journalisation d’une cellule</span><span class="sxs-lookup"><span data-stu-id="230df-114">Read and log one cell</span></span>

<span data-ttu-id="230df-115">Cet exemple lit la valeur de **a1** et l’imprime sur la console.</span><span class="sxs-lookup"><span data-stu-id="230df-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="230df-116">Utiliser des dates</span><span class="sxs-lookup"><span data-stu-id="230df-116">Work with dates</span></span>

<span data-ttu-id="230df-117">Les exemples de cette section indiquent comment utiliser l’objet [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript.</span><span class="sxs-lookup"><span data-stu-id="230df-117">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="230df-118">L’exemple suivant obtient la date et l’heure actuelles, puis écrit ces valeurs dans deux cellules de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="230df-118">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="230df-119">L’exemple suivant lit une date stockée dans Excel et la convertit en un objet JavaScript date.</span><span class="sxs-lookup"><span data-stu-id="230df-119">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="230df-120">Il utilise le [numéro de série numérique de la date](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) comme entrée pour la date JavaScript.</span><span class="sxs-lookup"><span data-stu-id="230df-120">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Read a date at cell A1 from Excel.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  dateRange.load("values");
  await context.sync();

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.values[0][0];
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="230df-121">Afficher les données</span><span class="sxs-lookup"><span data-stu-id="230df-121">Display data</span></span>

<span data-ttu-id="230df-122">Ces exemples montrent comment utiliser les données de feuille de calcul et fournir aux utilisateurs une meilleure vue ou organisation.</span><span class="sxs-lookup"><span data-stu-id="230df-122">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="230df-123">Application d’une mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="230df-123">Apply conditional formatting</span></span>

<span data-ttu-id="230df-124">Cet exemple applique la mise en forme conditionnelle à la plage utilisée dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="230df-124">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="230df-125">La mise en forme conditionnelle est un remplissage vert pour les 10% de valeurs les plus fréquentes.</span><span class="sxs-lookup"><span data-stu-id="230df-125">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="230df-126">Créer un tableau trié</span><span class="sxs-lookup"><span data-stu-id="230df-126">Create a sorted table</span></span>

<span data-ttu-id="230df-127">Cet exemple montre comment créer un tableau à partir de la plage utilisée dans la feuille de calcul active, puis comment le trier en fonction de la première colonne.</span><span class="sxs-lookup"><span data-stu-id="230df-127">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

## <a name="collaboration"></a><span data-ttu-id="230df-128">Collaboration</span><span class="sxs-lookup"><span data-stu-id="230df-128">Collaboration</span></span>

<span data-ttu-id="230df-129">Ces exemples montrent comment utiliser les fonctionnalités liées à la collaboration d’Excel, telles que les commentaires.</span><span class="sxs-lookup"><span data-stu-id="230df-129">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="230df-130">Supprimer les commentaires résolus</span><span class="sxs-lookup"><span data-stu-id="230df-130">Delete resolved comments</span></span>

<span data-ttu-id="230df-131">Cet exemple montre comment supprimer tous les commentaires résolus de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="230df-131">This sample deletes all resolved comments from the current worksheet.</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="230df-132">Exemples de scénario</span><span class="sxs-lookup"><span data-stu-id="230df-132">Scenario samples</span></span>

<span data-ttu-id="230df-133">Pour obtenir des exemples illustrant des solutions plus étendues dans le monde réel, consultez [exemples de scénarios pour les scripts Office](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="230df-133">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="230df-134">Suggérer de nouveaux exemples</span><span class="sxs-lookup"><span data-stu-id="230df-134">Suggest new samples</span></span>

<span data-ttu-id="230df-135">Nous vous invitons à suggérer de nouveaux exemples.</span><span class="sxs-lookup"><span data-stu-id="230df-135">We welcome suggestions for new samples.</span></span> <span data-ttu-id="230df-136">S’il existe un scénario courant qui aide les autres développeurs de script, veuillez nous en indiquer dans la section commentaires ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="230df-136">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
