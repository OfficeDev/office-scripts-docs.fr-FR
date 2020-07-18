---
title: Appeler des scripts à partir d’un flux manuel Power Automate
description: Un tutoriel sur l’utilisation des scripts Office dans Power Automate via un déclencheur manuel.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160434"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="acb93-103">Appeler des scripts à partir d’un flux manuel Power Automate (préversion)</span><span class="sxs-lookup"><span data-stu-id="acb93-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="acb93-104">Ce tutoriel vous apprend à exécuter un script Office pour Excel sur le web via [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="acb93-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="acb93-105">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="acb93-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="acb93-106">Ce tutoriel suppose que vous avez terminé le tutoriel [Enregistrer, modifier et créer des scripts Office dans Excel sur le web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="acb93-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="acb93-107">Préparer le classeur</span><span class="sxs-lookup"><span data-stu-id="acb93-107">Prepare the workbook</span></span>

<span data-ttu-id="acb93-108">Power Automate ne peut pas utiliser de références relatives comme `Workbook.getActiveWorksheet` pour accéder aux composants du classeur.</span><span class="sxs-lookup"><span data-stu-id="acb93-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="acb93-109">Nous avons donc besoin d’un classeur et d’une feuille de calcul avec des noms cohérents que Power Automate peut référencer.</span><span class="sxs-lookup"><span data-stu-id="acb93-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="acb93-110">Créer un classeur nommé **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="acb93-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="acb93-111">Dans le classeur **MyWorkbook**, créez une feuille de calcul appelée **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="acb93-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="acb93-112">Créer un script Office</span><span class="sxs-lookup"><span data-stu-id="acb93-112">Create an Office Script</span></span>

1. <span data-ttu-id="acb93-113">Accédez à l’onglet **Automatiser**, puis sélectionnez **Éditeur de code**.</span><span class="sxs-lookup"><span data-stu-id="acb93-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="acb93-114">Sélectionnez **Nouveau script**.</span><span class="sxs-lookup"><span data-stu-id="acb93-114">Select **New Script**.</span></span>

3. <span data-ttu-id="acb93-115">Remplacez le script par défaut par le script suivant.</span><span class="sxs-lookup"><span data-stu-id="acb93-115">Replace the default script with the following script.</span></span> <span data-ttu-id="acb93-116">Ce script ajoute la date et l’heure actuelles aux deux premières cellules de la feuille de calcul **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="acb93-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. <span data-ttu-id="acb93-117">Renommez le script **Définir la date et l’heure**.</span><span class="sxs-lookup"><span data-stu-id="acb93-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="acb93-118">Appuyez sur le nom du script pour le changer.</span><span class="sxs-lookup"><span data-stu-id="acb93-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="acb93-119">Enregistrez le script en appuyant sur **Enregistrer le script**.</span><span class="sxs-lookup"><span data-stu-id="acb93-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="acb93-120">Créer un flux de travail automatisé avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="acb93-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="acb93-121">Connectez-vous au site [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="acb93-121">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="acb93-122">Dans le menu qui s’affiche sur le côté gauche de l’écran, appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="acb93-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="acb93-123">Cela affiche une liste des moyens de créer de nouveaux flux de travail.</span><span class="sxs-lookup"><span data-stu-id="acb93-123">This brings you to list of ways to create new workflows.</span></span>

    ![Le bouton Créer dans Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="acb93-125">Dans la section **Démarrer à partir de zéro**, sélectionnez **Flux instantané**.</span><span class="sxs-lookup"><span data-stu-id="acb93-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="acb93-126">Cela crée un flux de travail activé manuellement.</span><span class="sxs-lookup"><span data-stu-id="acb93-126">This creates a manually activated workflow.</span></span>

    ![L’option Flux instantané pour créer un nouveau flux de travail.](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="acb93-128">Dans la boîte de dialogue qui apparaît, entrez un nom pour votre flux dans la zone de texte **Nom du flux**, sélectionnez **Déclencher manuellement un flux** dans la liste des options sous **Choisir le déclencheur du flux**, puis appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="acb93-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![L’option de déclenchement manuel pour créer un nouveau flux instantané.](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="acb93-130">Notez qu’un flux déclenché manuellement n’est que l’un des nombreux types de flux.</span><span class="sxs-lookup"><span data-stu-id="acb93-130">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="acb93-131">Dans le tutoriel suivant, vous allez créer un flux qui s’exécute automatiquement lorsque vous recevez un e-mail.</span><span class="sxs-lookup"><span data-stu-id="acb93-131">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="acb93-132">Appuyez sur **Nouvelle étape**.</span><span class="sxs-lookup"><span data-stu-id="acb93-132">Press **New step**.</span></span>

6. <span data-ttu-id="acb93-133">Sélectionnez l’onglet **Standard**, puis sélectionnez **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="acb93-133">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![L’option Power Automate pour Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="acb93-135">Sous **Actions**, sélectionnez **Exécuter le script** (préversion).</span><span class="sxs-lookup"><span data-stu-id="acb93-135">Under **Actions**, select **Run script (preview)**.</span></span>

    ![L’option d’action Power Automate pour exécuter le script (préversion).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="acb93-137">Spécifiez les paramètres suivants pour le connecteur **Exécuter le script** :</span><span class="sxs-lookup"><span data-stu-id="acb93-137">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="acb93-138">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="acb93-138">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="acb93-139">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="acb93-139">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="acb93-140">**Fichier** : MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="acb93-140">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="acb93-141">**Script** : Définir la date et l’heure</span><span class="sxs-lookup"><span data-stu-id="acb93-141">**Script**: Set date and time</span></span>

    ![Les paramètres du connecteur pour exécuter un script dans Power Automate.](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="acb93-143">Appuyez sur **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="acb93-143">Press **Save**.</span></span>

<span data-ttu-id="acb93-144">Votre flux est maintenant prêt à être exécuté via Power Automate.</span><span class="sxs-lookup"><span data-stu-id="acb93-144">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="acb93-145">Vous pouvez le tester à l’aide du bouton **Tester** dans l’éditeur de flux ou suivre les étapes restantes du tutoriel pour exécuter le flux à partir de votre collection de flux.</span><span class="sxs-lookup"><span data-stu-id="acb93-145">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="acb93-146">Exécuter le script via Power Automate</span><span class="sxs-lookup"><span data-stu-id="acb93-146">Run the script through Power Automate</span></span>

1. <span data-ttu-id="acb93-147">Sur la page principale de Power Automate, sélectionnez **Mes flux**.</span><span class="sxs-lookup"><span data-stu-id="acb93-147">From the main Power Automate page, select **My flows**.</span></span>

    ![Le bouton Mes flux dans Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="acb93-149">Sélectionnez **Mon flux de tutoriel** dans la liste des flux affichée dans l’onglet **Mes flux**. Cela affiche les informations sur le flux que nous avons créé précédemment.</span><span class="sxs-lookup"><span data-stu-id="acb93-149">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="acb93-150">Appuyez sur **Exécuter**.</span><span class="sxs-lookup"><span data-stu-id="acb93-150">Press **Run**.</span></span>

    ![Le bouton Exécuter dans Power Automate.](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="acb93-152">Un volet des tâches apparaîtra pour exécuter le flux.</span><span class="sxs-lookup"><span data-stu-id="acb93-152">A task pane will appear for running the flow.</span></span> <span data-ttu-id="acb93-153">Si vous êtes invité à vous **Connecter** à Excel Online, faites-le en appuyant sur **Continuer**.</span><span class="sxs-lookup"><span data-stu-id="acb93-153">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="acb93-154">Appuyez sur **Exécuter le flux**.</span><span class="sxs-lookup"><span data-stu-id="acb93-154">Press **Run flow**.</span></span> <span data-ttu-id="acb93-155">Cela exécute le flux, qui exécute le script Office associé.</span><span class="sxs-lookup"><span data-stu-id="acb93-155">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="acb93-156">Appuyez sur **Terminé**.</span><span class="sxs-lookup"><span data-stu-id="acb93-156">Press **Done**.</span></span> <span data-ttu-id="acb93-157">Vous devriez voir la section **Exécutions** s’actualiser en conséquence.</span><span class="sxs-lookup"><span data-stu-id="acb93-157">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="acb93-158">Actualisez la page pour voir les résultats de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="acb93-158">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="acb93-159">Si l’opération est réussie, accédez au classeur pour voir les cellules mises à jour.</span><span class="sxs-lookup"><span data-stu-id="acb93-159">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="acb93-160">Si l’opération a échoué, vérifiez les paramètres du flux et exécutez-le une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="acb93-160">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Production Power Automate indiquant une exécution de flux réussie.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="acb93-162">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="acb93-162">Next steps</span></span>

<span data-ttu-id="acb93-163">Suivez le tutoriel [Transférer des données aux scripts dans un flux Power Automate exécuté automatiquement](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="acb93-163">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="acb93-164">Il vous explique comment transmettre les données d’un service de flux de travail à votre script Office et comment exécuter le flux Power Automate lorsque certains événements se produisent.</span><span class="sxs-lookup"><span data-stu-id="acb93-164">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
