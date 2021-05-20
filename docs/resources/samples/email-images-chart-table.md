---
title: Envoyez par courriel les images d’un Excel graphique et d’une table
description: Apprenez à utiliser les scripts Office les Power Automate pour extraire et envoyer par courriel les images d’un graphique et d Excel table.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545774"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="e69cf-103">Utilisez Office scripts et des Power Automate pour envoyer des images d’un graphique et d’une table</span><span class="sxs-lookup"><span data-stu-id="e69cf-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="e69cf-104">Cet exemple utilise Office scripts et Power Automate pour créer un graphique.</span><span class="sxs-lookup"><span data-stu-id="e69cf-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="e69cf-105">Il envoie ensuite des images du graphique et de sa table de base.</span><span class="sxs-lookup"><span data-stu-id="e69cf-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="e69cf-106">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="e69cf-106">Example scenario</span></span>

* <span data-ttu-id="e69cf-107">Calculez pour obtenir les derniers résultats.</span><span class="sxs-lookup"><span data-stu-id="e69cf-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="e69cf-108">Créer un graphique.</span><span class="sxs-lookup"><span data-stu-id="e69cf-108">Create chart.</span></span>
* <span data-ttu-id="e69cf-109">Obtenez des images de tableau et de table.</span><span class="sxs-lookup"><span data-stu-id="e69cf-109">Get chart and table images.</span></span>
* <span data-ttu-id="e69cf-110">Envoyez les images par courriel Power Automate.</span><span class="sxs-lookup"><span data-stu-id="e69cf-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="e69cf-111">_Données d’entrée_</span><span class="sxs-lookup"><span data-stu-id="e69cf-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Une feuille de travail montrant un tableau des données d’entrée":::

<span data-ttu-id="e69cf-113">_Graphique de sortie_</span><span class="sxs-lookup"><span data-stu-id="e69cf-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Le graphique de colonne créé montrant le montant dû par le client":::

<span data-ttu-id="e69cf-115">_Courriel qui a été reçu par Power Automate flux_</span><span class="sxs-lookup"><span data-stu-id="e69cf-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="L’e-mail envoyé par le flux montrant Excel graphique intégré dans le corps":::

## <a name="solution"></a><span data-ttu-id="e69cf-117">Solution</span><span class="sxs-lookup"><span data-stu-id="e69cf-117">Solution</span></span>

<span data-ttu-id="e69cf-118">Cette solution a deux parties :</span><span class="sxs-lookup"><span data-stu-id="e69cf-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="e69cf-119">Un script Office pour calculer et extraire le graphique Excel la table</span><span class="sxs-lookup"><span data-stu-id="e69cf-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="e69cf-120">Un flux Power Automate pour invoquer le script et envoyer les résultats par courriel.</span><span class="sxs-lookup"><span data-stu-id="e69cf-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="e69cf-121">Par exemple, sur la façon de le faire, voir [Créer un flux de travail automatisé avec Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="e69cf-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="e69cf-122">Code de l’échantillon : Calculer et extraire Excel graphique et la table</span><span class="sxs-lookup"><span data-stu-id="e69cf-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="e69cf-123">Le script suivant calcule et extrait un graphique Excel tableau et une table.</span><span class="sxs-lookup"><span data-stu-id="e69cf-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="e69cf-124">Téléchargez l’exemple <a href="email-chart-table.xlsx"> deemail-chart-table.xlsxet </a> utilisez-le avec ce script pour l’essayer vous-même!</span><span class="sxs-lookup"><span data-stu-id="e69cf-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="e69cf-125">Power Automate : Envoyez un e-mail au graphique et aux images de table</span><span class="sxs-lookup"><span data-stu-id="e69cf-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="e69cf-126">Ce flux exécute le script et envoie les e-mails les images retournées.</span><span class="sxs-lookup"><span data-stu-id="e69cf-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="e69cf-127">Créez un nouveau **flux cloud instantané**.</span><span class="sxs-lookup"><span data-stu-id="e69cf-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="e69cf-128">Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="e69cf-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="e69cf-129">Ajoutez une **nouvelle étape qui** utilise le **connecteur Excel en ligne (Business)** avec l’action **de script Run.**</span><span class="sxs-lookup"><span data-stu-id="e69cf-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="e69cf-130">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="e69cf-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="e69cf-131">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="e69cf-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="e69cf-132">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="e69cf-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="e69cf-133">**Fichier**: Votre cahier de travail [(sélectionné avec le choix du fichier)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="e69cf-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="e69cf-134">**Script**: Votre nom de script</span><span class="sxs-lookup"><span data-stu-id="e69cf-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Le connecteur Excel en ligne (Business) terminé en Power Automate":::
1. <span data-ttu-id="e69cf-136">Cet exemple utilise Outlook client de messagerie.</span><span class="sxs-lookup"><span data-stu-id="e69cf-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="e69cf-137">Vous pouvez utiliser n’importe quel connecteur de messagerie Power Automate supports, mais le reste des étapes supposent que vous avez choisi Outlook.</span><span class="sxs-lookup"><span data-stu-id="e69cf-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="e69cf-138">Ajoutez une **nouvelle étape qui** utilise le **connecteur Office 365 Outlook et** l’action Envoyer et envoyer des **e-mails (V2).**</span><span class="sxs-lookup"><span data-stu-id="e69cf-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="e69cf-139">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="e69cf-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="e69cf-140">**À**: Votre compte de messagerie de test (ou e-mail personnel)</span><span class="sxs-lookup"><span data-stu-id="e69cf-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="e69cf-141">**Objet**: Veuillez examiner les données du rapport</span><span class="sxs-lookup"><span data-stu-id="e69cf-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="e69cf-142">Pour le **champ** Corps, sélectionnez « Code View » `</>` () et entrez ce qui suit:</span><span class="sxs-lookup"><span data-stu-id="e69cf-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Le connecteur Office 365 Outlook terminé en Power Automate":::
1. <span data-ttu-id="e69cf-144">Enregistrez le flux et essayez-le.</span><span class="sxs-lookup"><span data-stu-id="e69cf-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="e69cf-145">Vidéo de formation : Extraire et envoyer par courriel des images du graphique et de la table</span><span class="sxs-lookup"><span data-stu-id="e69cf-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="e69cf-146">[Regardez Sudhi Ramamurthy marcher à travers cet échantillon sur YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="e69cf-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
