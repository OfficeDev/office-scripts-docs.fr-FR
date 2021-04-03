---
title: Envoyer par courrier électronique les images d’un graphique et d’un tableau Excel
description: Découvrez comment utiliser Office Scripts et Power Automate pour extraire et envoyer par courrier électronique les images d’un graphique et d’un tableau Excel.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 7eb12526f97d72de31acdc3c9a4228c670875e2b
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571192"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="1e705-103">Utiliser Office Scripts et Power Automate pour envoyer des images électroniques d’un graphique et d’un tableau</span><span class="sxs-lookup"><span data-stu-id="1e705-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="1e705-104">Cet exemple utilise Office Scripts et Power Automate pour créer un graphique.</span><span class="sxs-lookup"><span data-stu-id="1e705-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="1e705-105">Il envoie ensuite des images du graphique et de sa table de base par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="1e705-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="1e705-106">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="1e705-106">Example scenario</span></span>

* <span data-ttu-id="1e705-107">Calculer pour obtenir les derniers résultats.</span><span class="sxs-lookup"><span data-stu-id="1e705-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="1e705-108">Créez un graphique.</span><span class="sxs-lookup"><span data-stu-id="1e705-108">Create chart.</span></span>
* <span data-ttu-id="1e705-109">Obtenir des images de graphique et de tableau.</span><span class="sxs-lookup"><span data-stu-id="1e705-109">Get chart and table images.</span></span>
* <span data-ttu-id="1e705-110">Envoyez un e-mail aux images avec Power Automate.</span><span class="sxs-lookup"><span data-stu-id="1e705-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="1e705-111">_Données d’entrée_</span><span class="sxs-lookup"><span data-stu-id="1e705-111">_Input data_</span></span>

![Données d’entrée](../../images/input-data.png)

<span data-ttu-id="1e705-113">_Graphique de sortie_</span><span class="sxs-lookup"><span data-stu-id="1e705-113">_Output chart_</span></span>

![Graphique créé](../../images/chart-created.png)

<span data-ttu-id="1e705-115">_Courrier électronique reçu via le flux Power Automate_</span><span class="sxs-lookup"><span data-stu-id="1e705-115">_Email that was received through Power Automate flow_</span></span>

![Courrier électronique reçu](../../images/email-received.png)

## <a name="solution"></a><span data-ttu-id="1e705-117">Solution</span><span class="sxs-lookup"><span data-stu-id="1e705-117">Solution</span></span>

<span data-ttu-id="1e705-118">Cette solution est en deux parties :</span><span class="sxs-lookup"><span data-stu-id="1e705-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="1e705-119">Script Office pour calculer et extraire un graphique et un tableau Excel</span><span class="sxs-lookup"><span data-stu-id="1e705-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="1e705-120">Flux Power Automate pour appeler le script et envoyer par courrier électronique les résultats.</span><span class="sxs-lookup"><span data-stu-id="1e705-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="1e705-121">Pour obtenir un exemple sur la procédure à suivre, voir Créer un flux de travail [automatisé avec Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="1e705-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="1e705-122">Exemple de code : calculer et extraire un graphique et un tableau Excel</span><span class="sxs-lookup"><span data-stu-id="1e705-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="1e705-123">Le script suivant calcule et extrait un graphique et un tableau Excel.</span><span class="sxs-lookup"><span data-stu-id="1e705-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="1e705-124">Téléchargez l’exemple <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> fichier et utilisez-le avec ce script pour l’essayer vous-même !</span><span class="sxs-lookup"><span data-stu-id="1e705-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="1e705-125">Vidéo de formation : extraire et envoyer des images par courrier électronique à un graphique et un tableau</span><span class="sxs-lookup"><span data-stu-id="1e705-125">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="1e705-126">[![Regardez une vidéo pas à pas sur l’extraction et l’envoi par courrier électronique d’images de graphique et de tableau](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vidéo pas à pas sur l’extraction et l’envoi par courrier électronique d’images de graphique et de tableau")</span><span class="sxs-lookup"><span data-stu-id="1e705-126">[![Watch step-by-step video on how to extract and email images of chart and table](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Step-by-step video on how to extract and email images of chart and table")</span></span>
