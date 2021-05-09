---
title: Exécuter un script sur tous les fichiers Excel d’un dossier
description: Découvrez comment exécuter un script sur tous les fichiers Excel dans un dossier sur OneDrive Entreprise.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: cfe603f3b7fa0ffc27aa3478b2f54788ad645b3f
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285807"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="35e44-103">Exécuter un script sur tous les fichiers Excel d’un dossier</span><span class="sxs-lookup"><span data-stu-id="35e44-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="35e44-104">Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise.</span><span class="sxs-lookup"><span data-stu-id="35e44-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="35e44-105">Il peut également être utilisé sur un SharePoint dossier.</span><span class="sxs-lookup"><span data-stu-id="35e44-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="35e44-106">Il effectue des calculs sur les fichiers Excel, ajoute une mise en forme et insère un [commentaire](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) qui @mentions collègue.</span><span class="sxs-lookup"><span data-stu-id="35e44-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="35e44-107">Téléchargez le fichier <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip,</a>extrayez les fichiers dans un dossier intitulé **Ventes** utilisés dans cet exemple et essayez-le vous-même !</span><span class="sxs-lookup"><span data-stu-id="35e44-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="35e44-108">Exemple de code : ajouter une mise en forme et insérer un commentaire</span><span class="sxs-lookup"><span data-stu-id="35e44-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="35e44-109">Il s’agit du script qui s’exécute sur chaque workbook individuel.</span><span class="sxs-lookup"><span data-stu-id="35e44-109">This is the script that runs on each individual workbook.</span></span>

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="35e44-110">Power Automate flux : exécuter le script sur chaque classeur du dossier</span><span class="sxs-lookup"><span data-stu-id="35e44-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="35e44-111">Ce flux exécute le script sur chaque classeur dans le dossier « Ventes ».</span><span class="sxs-lookup"><span data-stu-id="35e44-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="35e44-112">Créez un **flux de cloud instantané.**</span><span class="sxs-lookup"><span data-stu-id="35e44-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="35e44-113">Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**</span><span class="sxs-lookup"><span data-stu-id="35e44-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="35e44-114">Ajoutez **une nouvelle étape qui** utilise le connecteur **OneDrive Entreprise** et les fichiers de liste **dans l’action de** dossier.</span><span class="sxs-lookup"><span data-stu-id="35e44-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate":::
1. <span data-ttu-id="35e44-116">Sélectionnez le dossier « Ventes » avec les classeurs extraits.</span><span class="sxs-lookup"><span data-stu-id="35e44-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="35e44-117">Pour vous assurer que seuls les workbooks sont sélectionnés, choisissez **Nouvelle étape,** puis **sélectionnez Condition** et définissez les valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="35e44-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="35e44-118">**Nom** (valeur OneDrive nom de fichier)</span><span class="sxs-lookup"><span data-stu-id="35e44-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="35e44-119">« se termine par »</span><span class="sxs-lookup"><span data-stu-id="35e44-119">"ends with"</span></span>
    1. <span data-ttu-id="35e44-120">« xlsx ».</span><span class="sxs-lookup"><span data-stu-id="35e44-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Le bloc Power Automate condition qui applique les actions suivantes à chaque fichier":::
1. <span data-ttu-id="35e44-122">Sous la **branche Si oui,** ajoutez **le connecteur Excel Online (Entreprise)** avec l’action **Exécuter le script (prévisualisation).**</span><span class="sxs-lookup"><span data-stu-id="35e44-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="35e44-123">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="35e44-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="35e44-124">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="35e44-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="35e44-125">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="35e44-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="35e44-126">**Fichier**: **ID** (valeur OneDrive’ID de fichier)</span><span class="sxs-lookup"><span data-stu-id="35e44-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="35e44-127">**Script**: nom de votre script</span><span class="sxs-lookup"><span data-stu-id="35e44-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Le connecteur Excel Online (Entreprise) terminé dans Power Automate":::
1. <span data-ttu-id="35e44-129">Enregistrez le flux et testez-le.</span><span class="sxs-lookup"><span data-stu-id="35e44-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="35e44-130">Vidéo de formation : exécuter un script sur tous Excel fichiers d’un dossier</span><span class="sxs-lookup"><span data-stu-id="35e44-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="35e44-131">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="35e44-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
