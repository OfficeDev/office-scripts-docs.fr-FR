---
title: Envoyer par e-mail les images d Excel graphique et d’un tableau
description: Découvrez comment utiliser Office scripts et Power Automate pour extraire et envoyer par e-mail les images d’un Excel graphique et d’un tableau.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545774"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="7ed58-103">Utiliser Office scripts et Power Automate pour envoyer des images électroniques d’un graphique et d’un tableau</span><span class="sxs-lookup"><span data-stu-id="7ed58-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="7ed58-104">Cet exemple utilise Office scripts et Power Automate pour créer un graphique.</span><span class="sxs-lookup"><span data-stu-id="7ed58-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="7ed58-105">Il envoie ensuite des images du graphique et de sa table de base par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="7ed58-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="7ed58-106">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="7ed58-106">Example scenario</span></span>

* <span data-ttu-id="7ed58-107">Calculer pour obtenir les derniers résultats.</span><span class="sxs-lookup"><span data-stu-id="7ed58-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="7ed58-108">Créez un graphique.</span><span class="sxs-lookup"><span data-stu-id="7ed58-108">Create chart.</span></span>
* <span data-ttu-id="7ed58-109">Obtenir des images de graphique et de tableau.</span><span class="sxs-lookup"><span data-stu-id="7ed58-109">Get chart and table images.</span></span>
* <span data-ttu-id="7ed58-110">Envoyez un e-mail à l’Power Automate.</span><span class="sxs-lookup"><span data-stu-id="7ed58-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="7ed58-111">_Données d’entrée_</span><span class="sxs-lookup"><span data-stu-id="7ed58-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Feuille de calcul montrant une table des données d’entrée":::

<span data-ttu-id="7ed58-113">_Graphique de sortie_</span><span class="sxs-lookup"><span data-stu-id="7ed58-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Graphique en colonnes créé montrant le montant dû par le client":::

<span data-ttu-id="7ed58-115">_Courrier électronique reçu par le biais Power Automate flux_</span><span class="sxs-lookup"><span data-stu-id="7ed58-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Courrier électronique envoyé par le flux montrant le graphique Excel incorporé dans le corps":::

## <a name="solution"></a><span data-ttu-id="7ed58-117">Solution</span><span class="sxs-lookup"><span data-stu-id="7ed58-117">Solution</span></span>

<span data-ttu-id="7ed58-118">Cette solution est en deux parties :</span><span class="sxs-lookup"><span data-stu-id="7ed58-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="7ed58-119">Un script Office pour calculer et extraire Excel graphique et tableau</span><span class="sxs-lookup"><span data-stu-id="7ed58-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="7ed58-120">Un flux Power Automate pour appeler le script et envoyer par courrier électronique les résultats.</span><span class="sxs-lookup"><span data-stu-id="7ed58-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="7ed58-121">Pour obtenir un exemple sur la procédure à suivre, voir Créer un flux de travail automatisé [avec Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="7ed58-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="7ed58-122">Exemple de code : calculer et extraire Excel graphique et tableau</span><span class="sxs-lookup"><span data-stu-id="7ed58-122">Sample code: Calculate and extract Excel chart and table</span></span>

<span data-ttu-id="7ed58-123">Le script suivant calcule et extrait un Excel graphique et un tableau.</span><span class="sxs-lookup"><span data-stu-id="7ed58-123">The following script calculates and extracts an Excel chart and table.</span></span>

<span data-ttu-id="7ed58-124">Téléchargez l’exemple <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> fichier et utilisez-le avec ce script pour l’essayer vous-même !</span><span class="sxs-lookup"><span data-stu-id="7ed58-124">Download the sample file <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="7ed58-125">Power Automate flux : envoyer un e-mail aux images du graphique et du tableau</span><span class="sxs-lookup"><span data-stu-id="7ed58-125">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="7ed58-126">Ce flux exécute le script et envoie par e-mail les images renvoyées.</span><span class="sxs-lookup"><span data-stu-id="7ed58-126">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="7ed58-127">Créez un **flux de cloud instantané.**</span><span class="sxs-lookup"><span data-stu-id="7ed58-127">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="7ed58-128">Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**</span><span class="sxs-lookup"><span data-stu-id="7ed58-128">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="7ed58-129">Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** avec l’action **de script Exécuter.**</span><span class="sxs-lookup"><span data-stu-id="7ed58-129">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="7ed58-130">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="7ed58-130">Use the following values for the action:</span></span>
    * <span data-ttu-id="7ed58-131">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="7ed58-131">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="7ed58-132">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="7ed58-132">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="7ed58-133">**Fichier**: votre classeur [(sélectionné avec le sélecateur de fichiers)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="7ed58-133">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="7ed58-134">**Script**: nom de votre script</span><span class="sxs-lookup"><span data-stu-id="7ed58-134">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Le connecteur Excel Online (Entreprise) terminé dans Power Automate":::
1. <span data-ttu-id="7ed58-136">Cet exemple utilise Outlook client de messagerie.</span><span class="sxs-lookup"><span data-stu-id="7ed58-136">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="7ed58-137">Vous pouvez utiliser n’importe quel connecteur de messagerie Power Automate prend en charge, mais le reste des étapes suppose que vous avez choisi Outlook.</span><span class="sxs-lookup"><span data-stu-id="7ed58-137">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="7ed58-138">Ajoutez **une nouvelle étape** qui utilise le connecteur **Office 365 Outlook** et l’action Envoyer et e-mail **(V2).**</span><span class="sxs-lookup"><span data-stu-id="7ed58-138">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="7ed58-139">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="7ed58-139">Use the following values for the action:</span></span>
    * <span data-ttu-id="7ed58-140">**À**: Votre compte de messagerie de test (ou e-mail personnel)</span><span class="sxs-lookup"><span data-stu-id="7ed58-140">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="7ed58-141">**Objet :** Veuillez consulter les données du rapport</span><span class="sxs-lookup"><span data-stu-id="7ed58-141">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="7ed58-142">Pour le **champ Corps,** sélectionnez « Affichage de code » ( `</>` ), puis entrez les entrées suivantes :</span><span class="sxs-lookup"><span data-stu-id="7ed58-142">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Connecteur de Office 365 Outlook terminé dans Power Automate":::
1. <span data-ttu-id="7ed58-144">Enregistrez le flux et testez-le.</span><span class="sxs-lookup"><span data-stu-id="7ed58-144">Save the flow and try it out.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="7ed58-145">Vidéo de formation : extraire et envoyer des images par courrier électronique à un graphique et un tableau</span><span class="sxs-lookup"><span data-stu-id="7ed58-145">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="7ed58-146">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="7ed58-146">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
