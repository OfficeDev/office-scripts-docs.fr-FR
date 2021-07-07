---
title: Exécuter un script sur tous les fichiers Excel d’un dossier
description: Découvrez comment exécuter un script sur tous les fichiers Excel dans un dossier sur OneDrive Entreprise.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: bf9c0c486dacced5c3017b267ea65dfd215a5197
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313896"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="22762-103">Exécuter un script sur tous les fichiers Excel d’un dossier</span><span class="sxs-lookup"><span data-stu-id="22762-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="22762-104">Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise.</span><span class="sxs-lookup"><span data-stu-id="22762-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="22762-105">Il peut également être utilisé sur un SharePoint dossier.</span><span class="sxs-lookup"><span data-stu-id="22762-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="22762-106">Il effectue des calculs sur les fichiers Excel, ajoute une mise en forme et insère un [commentaire](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) qui @mentions collègue.</span><span class="sxs-lookup"><span data-stu-id="22762-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="22762-107">Exemples Excel fichiers</span><span class="sxs-lookup"><span data-stu-id="22762-107">Sample Excel files</span></span>

<span data-ttu-id="22762-108">Téléchargez <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> tous les workbooks dont vous aurez besoin pour cet exemple.</span><span class="sxs-lookup"><span data-stu-id="22762-108">Download <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> for all the workbooks you'll need for this sample.</span></span> <span data-ttu-id="22762-109">Extrayer ces fichiers dans un dossier intitulé **Ventes**.</span><span class="sxs-lookup"><span data-stu-id="22762-109">Extract those files to a folder titled **Sales**.</span></span> <span data-ttu-id="22762-110">Ajoutez le script suivant à votre collection de scripts pour essayer l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="22762-110">Add the following script to your script collection to try the sample yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="22762-111">Exemple de code : ajouter une mise en forme et insérer un commentaire</span><span class="sxs-lookup"><span data-stu-id="22762-111">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="22762-112">Il s’agit du script qui s’exécute sur chaque workbook individuel.</span><span class="sxs-lookup"><span data-stu-id="22762-112">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="22762-113">Power Automate flux : exécuter le script sur chaque classeur du dossier</span><span class="sxs-lookup"><span data-stu-id="22762-113">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="22762-114">Ce flux exécute le script sur chaque classeur dans le dossier « Ventes ».</span><span class="sxs-lookup"><span data-stu-id="22762-114">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="22762-115">Créez un **flux de cloud instantané.**</span><span class="sxs-lookup"><span data-stu-id="22762-115">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="22762-116">Sélectionnez **Déclencher manuellement un flux,** puis **sélectionnez Créer.**</span><span class="sxs-lookup"><span data-stu-id="22762-116">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="22762-117">Ajoutez **une nouvelle étape qui** utilise le connecteur **OneDrive Entreprise** et les fichiers de liste **dans l’action de** dossier.</span><span class="sxs-lookup"><span data-stu-id="22762-117">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate.":::
1. <span data-ttu-id="22762-119">Sélectionnez le dossier « Ventes » avec les classeurs extraits.</span><span class="sxs-lookup"><span data-stu-id="22762-119">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="22762-120">Pour vous assurer que seuls les workbooks sont sélectionnés, choisissez **Nouvelle étape,** puis **sélectionnez Condition** et définissez les valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="22762-120">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="22762-121">**Nom** (valeur OneDrive nom de fichier)</span><span class="sxs-lookup"><span data-stu-id="22762-121">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="22762-122">« se termine par »</span><span class="sxs-lookup"><span data-stu-id="22762-122">"ends with"</span></span>
    1. <span data-ttu-id="22762-123">« xlsx ».</span><span class="sxs-lookup"><span data-stu-id="22762-123">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Le Power Automate condition qui applique les actions suivantes à chaque fichier.":::
1. <span data-ttu-id="22762-125">Sous la **branche Si oui,** ajoutez **le connecteur Excel Online (Entreprise)** avec l’action **de script Exécuter.**</span><span class="sxs-lookup"><span data-stu-id="22762-125">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="22762-126">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="22762-126">Use the following values for the action:</span></span>
    1. <span data-ttu-id="22762-127">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="22762-127">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="22762-128">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="22762-128">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="22762-129">**Fichier**: **ID** (valeur OneDrive’ID de fichier)</span><span class="sxs-lookup"><span data-stu-id="22762-129">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="22762-130">**Script**: nom de votre script</span><span class="sxs-lookup"><span data-stu-id="22762-130">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Le connecteur Excel Online (Entreprise) dans Power Automate.":::
1. <span data-ttu-id="22762-132">Enregistrez le flux et testez-le. Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.</span><span class="sxs-lookup"><span data-stu-id="22762-132">Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="22762-133">Vidéo de formation : exécuter un script sur tous Excel fichiers d’un dossier</span><span class="sxs-lookup"><span data-stu-id="22762-133">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="22762-134">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="22762-134">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
