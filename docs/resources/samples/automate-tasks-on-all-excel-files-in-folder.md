---
title: Exécuter un script sur tous les fichiers Excel d’un dossier
description: Découvrez comment exécuter un script sur tous les fichiers Excel dans un dossier sur OneDrive Entreprise.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a6b869e2b346635e2b28fa7c6273c1a86a5bc5c5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232626"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="b5451-103">Exécuter un script sur tous les fichiers Excel d’un dossier</span><span class="sxs-lookup"><span data-stu-id="b5451-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="b5451-104">Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise.</span><span class="sxs-lookup"><span data-stu-id="b5451-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="b5451-105">Il peut également être utilisé sur un SharePoint dossier.</span><span class="sxs-lookup"><span data-stu-id="b5451-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="b5451-106">Il effectue des calculs sur les fichiers Excel, ajoute une mise en forme et insère un [commentaire](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) qui @mentions collègue.</span><span class="sxs-lookup"><span data-stu-id="b5451-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

<span data-ttu-id="b5451-107">Téléchargez le fichier <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip,</a>extrayez les fichiers dans un dossier intitulé **Ventes** utilisés dans cet exemple, et essayez-le vous-même !</span><span class="sxs-lookup"><span data-stu-id="b5451-107">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="b5451-108">Exemple de code : ajouter une mise en forme et insérer un commentaire</span><span class="sxs-lookup"><span data-stu-id="b5451-108">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="b5451-109">Il s’agit du script qui s’exécute sur chaque workbook individuel.</span><span class="sxs-lookup"><span data-stu-id="b5451-109">This is the script that runs on each individual workbook.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a><span data-ttu-id="b5451-110">Power Automate flux : exécuter le script sur chaque classeur du dossier</span><span class="sxs-lookup"><span data-stu-id="b5451-110">Power Automate flow: Run the script on every workbook in the folder</span></span>

<span data-ttu-id="b5451-111">Ce flux exécute le script sur chaque classeur dans le dossier « Ventes ».</span><span class="sxs-lookup"><span data-stu-id="b5451-111">This flow runs the script on every workbook in the "Sales" folder.</span></span>

1. <span data-ttu-id="b5451-112">Créez un **flux de cloud instantané.**</span><span class="sxs-lookup"><span data-stu-id="b5451-112">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="b5451-113">Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**</span><span class="sxs-lookup"><span data-stu-id="b5451-113">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="b5451-114">Ajoutez **une nouvelle étape qui** utilise le connecteur **OneDrive Entreprise** et les fichiers de liste **dans l’action de** dossier.</span><span class="sxs-lookup"><span data-stu-id="b5451-114">Add a **New step** that uses the **OneDrive for Business** connector and the **List files in folder** action.</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate":::
1. <span data-ttu-id="b5451-116">Sélectionnez le dossier « Ventes » avec les classeurs extraits.</span><span class="sxs-lookup"><span data-stu-id="b5451-116">Select the "Sales" folder with the extracted workbooks.</span></span>
1. <span data-ttu-id="b5451-117">Pour vous assurer que seuls les workbooks sont sélectionnés, choisissez **Nouvelle étape,** puis **sélectionnez Condition** et définissez les valeurs suivantes :</span><span class="sxs-lookup"><span data-stu-id="b5451-117">To ensure only workbooks are selected, choose **New step**, then select **Condition** and set the following values:</span></span>
    1. <span data-ttu-id="b5451-118">**Nom** (valeur OneDrive nom de fichier)</span><span class="sxs-lookup"><span data-stu-id="b5451-118">**Name** (the OneDrive file name value)</span></span>
    1. <span data-ttu-id="b5451-119">« se termine par »</span><span class="sxs-lookup"><span data-stu-id="b5451-119">"ends with"</span></span>
    1. <span data-ttu-id="b5451-120">« xlsx ».</span><span class="sxs-lookup"><span data-stu-id="b5451-120">"xlsx".</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Le bloc Power Automate condition qui applique les actions suivantes à chaque fichier":::
1. <span data-ttu-id="b5451-122">Sous la **branche Si oui,** ajoutez **le connecteur Excel Online (Entreprise)** avec l’action **Exécuter le script (prévisualisation).**</span><span class="sxs-lookup"><span data-stu-id="b5451-122">Under the **If yes** branch, add the **Excel Online (Business)** connector with the **Run script (preview)** action.</span></span> <span data-ttu-id="b5451-123">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="b5451-123">Use the following values for the action:</span></span>
    1. <span data-ttu-id="b5451-124">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="b5451-124">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="b5451-125">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="b5451-125">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="b5451-126">**File**: **ID** (valeur OneDrive’ID de fichier)</span><span class="sxs-lookup"><span data-stu-id="b5451-126">**File**: **Id** (the OneDrive file ID value)</span></span>
    1. <span data-ttu-id="b5451-127">**Script**: nom de votre script</span><span class="sxs-lookup"><span data-stu-id="b5451-127">**Script**: Your script name</span></span>

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Le connecteur Excel Online (Entreprise) terminé dans Power Automate":::
1. <span data-ttu-id="b5451-129">Enregistrez le flux et testez-le.</span><span class="sxs-lookup"><span data-stu-id="b5451-129">Save the flow and try it out.</span></span>

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="b5451-130">Vidéo de formation : exécuter un script sur tous Excel fichiers d’un dossier</span><span class="sxs-lookup"><span data-stu-id="b5451-130">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="b5451-131">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/xMg711o7k6w).</span><span class="sxs-lookup"><span data-stu-id="b5451-131">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/xMg711o7k6w).</span></span>
