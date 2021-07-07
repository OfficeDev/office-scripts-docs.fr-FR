---
title: Envoyer par courrier électronique les images d Excel graphique et d’un tableau
description: Découvrez comment utiliser Office scripts et Power Automate pour extraire et envoyer par e-mail les images d’un Excel graphique et d’un tableau.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 50bc65c82df7f5fc68dbebf942c4f607bb6af60a
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313840"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a><span data-ttu-id="4f2ae-103">Utiliser Office scripts et Power Automate pour envoyer des images électroniques d’un graphique et d’un tableau</span><span class="sxs-lookup"><span data-stu-id="4f2ae-103">Use Office Scripts and Power Automate to email images of a chart and table</span></span>

<span data-ttu-id="4f2ae-104">Cet exemple utilise Office scripts et Power Automate pour créer un graphique.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-104">This sample uses Office Scripts and Power Automate to create a chart.</span></span> <span data-ttu-id="4f2ae-105">Il envoie ensuite des images du graphique et de sa table de base par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-105">It then emails images of the chart and its base table.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="4f2ae-106">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="4f2ae-106">Example scenario</span></span>

* <span data-ttu-id="4f2ae-107">Calculer pour obtenir les derniers résultats.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-107">Calculate to get latest results.</span></span>
* <span data-ttu-id="4f2ae-108">Créez un graphique.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-108">Create chart.</span></span>
* <span data-ttu-id="4f2ae-109">Obtenir des images de graphique et de tableau.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-109">Get chart and table images.</span></span>
* <span data-ttu-id="4f2ae-110">Envoyez un e-mail aux images Power Automate.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-110">Email the images with Power Automate.</span></span>

<span data-ttu-id="4f2ae-111">_Données d’entrée_</span><span class="sxs-lookup"><span data-stu-id="4f2ae-111">_Input data_</span></span>

:::image type="content" source="../../images/input-data.png" alt-text="Feuille de calcul montrant une table des données d’entrée.":::

<span data-ttu-id="4f2ae-113">_Graphique de sortie_</span><span class="sxs-lookup"><span data-stu-id="4f2ae-113">_Output chart_</span></span>

:::image type="content" source="../../images/chart-created.png" alt-text="Graphique en colonnes créé montrant le montant dû par le client.":::

<span data-ttu-id="4f2ae-115">_Courrier électronique reçu par le biais Power Automate flux_</span><span class="sxs-lookup"><span data-stu-id="4f2ae-115">_Email that was received through Power Automate flow_</span></span>

:::image type="content" source="../../images/email-received.png" alt-text="Courrier électronique envoyé par le flux montrant le Excel graphique incorporé dans le corps.":::

## <a name="solution"></a><span data-ttu-id="4f2ae-117">Solution</span><span class="sxs-lookup"><span data-stu-id="4f2ae-117">Solution</span></span>

<span data-ttu-id="4f2ae-118">Cette solution est en deux parties :</span><span class="sxs-lookup"><span data-stu-id="4f2ae-118">This solution has two parts:</span></span>

1. [<span data-ttu-id="4f2ae-119">Un script Office pour calculer et extraire Excel graphique et tableau</span><span class="sxs-lookup"><span data-stu-id="4f2ae-119">An Office Script to calculate and extract Excel chart and table</span></span>](#sample-code-calculate-and-extract-excel-chart-and-table)
1. <span data-ttu-id="4f2ae-120">Un flux Power Automate pour appeler le script et envoyer par courrier électronique les résultats.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-120">A Power Automate flow to invoke the script and email the results.</span></span> <span data-ttu-id="4f2ae-121">Pour obtenir un exemple sur la procédure à suivre, voir Créer un flux de travail automatisé [avec Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="4f2ae-121">For an example on how to do this, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="4f2ae-122">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="4f2ae-122">Sample Excel file</span></span>

<span data-ttu-id="4f2ae-123">Téléchargez <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> pour un livre de travail prêt à l’emploi.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-123">Download <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="4f2ae-124">Ajoutez le script suivant pour essayer l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="4f2ae-124">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a><span data-ttu-id="4f2ae-125">Exemple de code : calculer et extraire Excel graphique et tableau</span><span class="sxs-lookup"><span data-stu-id="4f2ae-125">Sample code: Calculate and extract Excel chart and table</span></span>

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a><span data-ttu-id="4f2ae-126">Power Automate flux : envoyer un e-mail aux images du graphique et du tableau</span><span class="sxs-lookup"><span data-stu-id="4f2ae-126">Power Automate flow: Email the chart and table images</span></span>

<span data-ttu-id="4f2ae-127">Ce flux exécute le script et envoie par e-mail les images renvoyées.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-127">This flow runs the script and emails the returned images.</span></span>

1. <span data-ttu-id="4f2ae-128">Créez un **flux de cloud instantané.**</span><span class="sxs-lookup"><span data-stu-id="4f2ae-128">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="4f2ae-129">Sélectionnez **Déclencher manuellement un flux,** puis **sélectionnez Créer.**</span><span class="sxs-lookup"><span data-stu-id="4f2ae-129">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="4f2ae-130">Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** avec l’action **exécuter le script.**</span><span class="sxs-lookup"><span data-stu-id="4f2ae-130">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="4f2ae-131">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="4f2ae-131">Use the following values for the action:</span></span>
    * <span data-ttu-id="4f2ae-132">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="4f2ae-132">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="4f2ae-133">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="4f2ae-133">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="4f2ae-134">**Fichier**: votre classeur [(sélectionné avec le sélecateur de fichiers)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)</span><span class="sxs-lookup"><span data-stu-id="4f2ae-134">**File**: Your workbook ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="4f2ae-135">**Script**: nom de votre script</span><span class="sxs-lookup"><span data-stu-id="4f2ae-135">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Le connecteur Excel Online (Entreprise) dans Power Automate.":::
1. <span data-ttu-id="4f2ae-137">Cet exemple utilise Outlook client de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-137">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="4f2ae-138">Vous pouvez utiliser n’importe quel connecteur de messagerie Power Automate prend en charge, mais le reste des étapes suppose que vous avez choisi Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-138">You could use any email connector Power Automate supports, but the rest of the steps assume that you chose Outlook.</span></span> <span data-ttu-id="4f2ae-139">Ajoutez **une nouvelle étape** qui utilise le connecteur **Office 365 Outlook** et l’action Envoyer et e-mail **(V2).**</span><span class="sxs-lookup"><span data-stu-id="4f2ae-139">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="4f2ae-140">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="4f2ae-140">Use the following values for the action:</span></span>
    * <span data-ttu-id="4f2ae-141">**À**: Votre compte de messagerie de test (ou e-mail personnel)</span><span class="sxs-lookup"><span data-stu-id="4f2ae-141">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="4f2ae-142">**Objet :** Veuillez consulter les données du rapport</span><span class="sxs-lookup"><span data-stu-id="4f2ae-142">**Subject**: Please Review Report Data</span></span>
    * <span data-ttu-id="4f2ae-143">Pour le **champ Corps,** sélectionnez « Affichage de code » ( `</>` ), puis entrez les entrées suivantes :</span><span class="sxs-lookup"><span data-stu-id="4f2ae-143">For the **Body** field, select "Code View" (`</>`) and enter the following:</span></span>

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Connecteur de Office 365 Outlook terminé dans Power Automate.":::
1. <span data-ttu-id="4f2ae-145">Enregistrez le flux et testez-le. Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.</span><span class="sxs-lookup"><span data-stu-id="4f2ae-145">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a><span data-ttu-id="4f2ae-146">Vidéo de formation : extraire et envoyer des images par courrier électronique à un graphique et un tableau</span><span class="sxs-lookup"><span data-stu-id="4f2ae-146">Training video: Extract and email images of chart and table</span></span>

<span data-ttu-id="4f2ae-147">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/152GJyqc-Kw).</span><span class="sxs-lookup"><span data-stu-id="4f2ae-147">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/152GJyqc-Kw).</span></span>
