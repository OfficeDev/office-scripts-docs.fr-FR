---
title: Exécuter un script sur tous les fichiers Excel d’un dossier
description: Découvrez comment exécuter un script sur tous les fichiers Excel dans un dossier sur OneDrive Entreprise.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545789"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="e46b7-103">Exécuter un script sur tous les fichiers Excel d’un dossier</span><span class="sxs-lookup"><span data-stu-id="e46b7-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="e46b7-104">Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise.</span><span class="sxs-lookup"><span data-stu-id="e46b7-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="e46b7-105">Il peut également être utilisé sur un dossier SharePoint dossier.</span><span class="sxs-lookup"><span data-stu-id="e46b7-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="e46b7-106">Il effectue des calculs sur les fichiers Excel, ajoute le formatage, et insère un [commentaire qui @mentions un](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) collègue.</span><span class="sxs-lookup"><span data-stu-id="e46b7-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="e46b7-107">Téléchargez le <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true"> fichierhighlight-alert-excel-files.zip</a>, extraire les fichiers dans un dossier intitulé **Ventes utilisées** dans cet échantillon, et l’essayer vous-même!</span><span class="sxs-lookup"><span data-stu-id="e46b7-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="e46b7-108">Exemple de code : Ajouter le formatage et insérer le commentaire</span><span class="sxs-lookup"><span data-stu-id="e46b7-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="e46b7-109">C’est le script qui s’exécute sur chaque cahier de travail individuel.</span><span class="sxs-lookup"><span data-stu-id="e46b7-109">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="e46b7-110">Power Automate flux : exécutez le script sur chaque cahier de travail dans le dossier</span><span class="sxs-lookup"><span data-stu-id="e46b7-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="e46b7-111">Ce flux exécute le script sur chaque cahier de travail dans le dossier « Ventes ».</span><span class="sxs-lookup"><span data-stu-id="e46b7-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="e46b7-112">Créez un nouveau **flux cloud instantané**.</span><span class="sxs-lookup"><span data-stu-id="e46b7-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="e46b7-113">Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="e46b7-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="e46b7-114">Ajoutez une **nouvelle étape qui** utilise le **connecteur OneDrive Entreprise** liste et les fichiers Liste dans **l’action du** dossier.</span><span class="sxs-lookup"><span data-stu-id="e46b7-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Le connecteur OneDrive Entreprise terminé en Power Automate":::
1. <span data-ttu-id="e46b7-116">Sélectionnez le dossier « Ventes » avec les cahiers de travail extraits.</span><span class="sxs-lookup"><span data-stu-id="e46b7-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="e46b7-117">Pour vous assurer que seuls les cahiers de travail sont sélectionnés, **choisissez Nouvelle étape,** puis **sélectionnez Condition** et définissez les valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="e46b7-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="e46b7-118">**Nom** (la valeur OneDrive nom du fichier)</span><span class="sxs-lookup"><span data-stu-id="e46b7-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="e46b7-119">« se termine par »</span><span class="sxs-lookup"><span data-stu-id="e46b7-119">"ends with"</span></span>
    1. <span data-ttu-id="e46b7-120">« xlsx ».</span><span class="sxs-lookup"><span data-stu-id="e46b7-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Le bloc Power Automate condition qui applique les actions ultérieures à chaque fichier":::
1. <span data-ttu-id="e46b7-122">Sous la **branche Si oui,** ajoutez le **connecteur Excel en ligne (Business)** avec l’action **script Run.**</span><span class="sxs-lookup"><span data-stu-id="e46b7-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="e46b7-123">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="e46b7-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="e46b7-124">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="e46b7-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="e46b7-125">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="e46b7-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="e46b7-126">**Fichier**: **Id** (la valeur d’identification OneDrive fichier)</span><span class="sxs-lookup"><span data-stu-id="e46b7-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="e46b7-127">**Script**: Votre nom de script</span><span class="sxs-lookup"><span data-stu-id="e46b7-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Le connecteur Excel en ligne (Business) terminé en Power Automate":::
1. <span data-ttu-id="e46b7-129">Enregistrez le flux et essayez-le.</span><span class="sxs-lookup"><span data-stu-id="e46b7-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="e46b7-130">Vidéo de formation : Exécutez un script sur tous les Excel fichiers dans un dossier</span><span class="sxs-lookup"><span data-stu-id="e46b7-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="e46b7-131">[Regardez Sudhi Ramamurthy marcher à travers cet échantillon sur YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="e46b7-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
