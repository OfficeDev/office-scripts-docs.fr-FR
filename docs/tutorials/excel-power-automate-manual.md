---
title: Appeler des scripts à partir d’un flux manuel Power Automate
description: Un tutoriel sur l’utilisation des scripts Office dans Power Automate via un déclencheur manuel.
ms.date: 12/28/2020
localization_priority: Priority
ms.openlocfilehash: 763a61f65afbdbb55346bb3d99f216da33e8419c
ms.sourcegitcommit: 9df67e007ddbfec79a7360df9f4ea5ac6c86fb08
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49772984"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="6836f-103">Appeler des scripts à partir d’un flux manuel Power Automate (préversion)</span><span class="sxs-lookup"><span data-stu-id="6836f-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="6836f-104">Ce tutoriel vous apprend à exécuter un script Office pour Excel sur le web via [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="6836f-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span> <span data-ttu-id="6836f-105">Vous allez créer un script qui met à jour les valeurs de deux cellules en y indiquant la date et l’heure de son exécution.</span><span class="sxs-lookup"><span data-stu-id="6836f-105">You'll make a script that updates the values of two cells with the current time.</span></span> <span data-ttu-id="6836f-106">Vous allez ensuite connecter ce script à un flux Power Automate déclenché manuellement, pour que le script s’exécute à chaque pression sur un bouton dans Power Automate.</span><span class="sxs-lookup"><span data-stu-id="6836f-106">You'll then connect that script to a manually triggered Power Automate flow, so that the script is run whenever a button in Power Automate is pressed.</span></span> <span data-ttu-id="6836f-107">Après avoir assimilé le modèle de base, vous pourrez développer le flux pour inclure d’autres applications et automatiser davantage votre flux de travail quotidien.</span><span class="sxs-lookup"><span data-stu-id="6836f-107">Once you understand the basic pattern, you can expand the flow to include other applications and automate more of your daily workflow.</span></span>

> [!TIP]
> <span data-ttu-id="6836f-108">Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="6836f-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="6836f-109">[La fonctionnalité Scripts Office utilise TypeScript](../overview/code-editor-environment.md), et ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="6836f-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="6836f-110">Si vous découvrez JavaScript, nous vous conseillons de commencer par consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="6836f-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6836f-111">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6836f-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="6836f-112">Préparer le classeur</span><span class="sxs-lookup"><span data-stu-id="6836f-112">Prepare the workbook</span></span>

<span data-ttu-id="6836f-113">Power Automate ne peut pas utiliser de [références relatives](../testing/power-automate-troubleshooting.md#avoid-using-relative-references) comme `Workbook.getActiveWorksheet`pour accéder aux composants du classeur.</span><span class="sxs-lookup"><span data-stu-id="6836f-113">Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="6836f-114">Nous avons donc besoin d’un classeur et d’une feuille de calcul avec des noms cohérents que Power Automate peut référencer.</span><span class="sxs-lookup"><span data-stu-id="6836f-114">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="6836f-115">Créer un classeur nommé **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="6836f-115">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="6836f-116">Dans le classeur **MyWorkbook**, créez une feuille de calcul appelée **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="6836f-116">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="6836f-117">Créer un script Office</span><span class="sxs-lookup"><span data-stu-id="6836f-117">Create an Office Script</span></span>

1. <span data-ttu-id="6836f-118">Accédez à l’onglet **Automatiser**, puis sélectionnez **Tous les scripts**.</span><span class="sxs-lookup"><span data-stu-id="6836f-118">Go to the **Automate** tab and select **All Scripts**.</span></span>

2. <span data-ttu-id="6836f-119">Sélectionnez **Nouveau script**.</span><span class="sxs-lookup"><span data-stu-id="6836f-119">Select **New Script**.</span></span>

3. <span data-ttu-id="6836f-120">Remplacez le script par défaut par le script suivant.</span><span class="sxs-lookup"><span data-stu-id="6836f-120">Replace the default script with the following script.</span></span> <span data-ttu-id="6836f-121">Ce script ajoute la date et l’heure actuelles aux deux premières cellules de la feuille de calcul **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="6836f-121">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

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

4. <span data-ttu-id="6836f-122">Renommez le script **Définir la date et l’heure**.</span><span class="sxs-lookup"><span data-stu-id="6836f-122">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="6836f-123">Appuyez sur le nom du script pour le changer.</span><span class="sxs-lookup"><span data-stu-id="6836f-123">Press the script name to change it.</span></span>

5. <span data-ttu-id="6836f-124">Enregistrez le script en appuyant sur **Enregistrer le script**.</span><span class="sxs-lookup"><span data-stu-id="6836f-124">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="6836f-125">Créer un flux de travail automatisé avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="6836f-125">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="6836f-126">Connectez-vous au site [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="6836f-126">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="6836f-127">Dans le menu qui s’affiche sur le côté gauche de l’écran, appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="6836f-127">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="6836f-128">Cela affiche une liste des moyens de créer de nouveaux flux de travail.</span><span class="sxs-lookup"><span data-stu-id="6836f-128">This brings you to list of ways to create new workflows.</span></span>

    ![Le bouton Créer dans Power Automate](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="6836f-130">Dans la section **Démarrer à partir de zéro**, sélectionnez **Flux instantané**.</span><span class="sxs-lookup"><span data-stu-id="6836f-130">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="6836f-131">Cela crée un flux de travail activé manuellement.</span><span class="sxs-lookup"><span data-stu-id="6836f-131">This creates a manually activated workflow.</span></span>

    ![L’option Flux instantané pour créer un nouveau flux de travail](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="6836f-133">Dans la boîte de dialogue qui apparaît, entrez un nom pour votre flux dans la zone de texte **Nom du flux**, sélectionnez **Déclencher manuellement un flux** dans la liste des options sous **Choisir le déclencheur du flux**, puis appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="6836f-133">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![L’option de déclenchement manuel pour créer un nouveau flux instantané](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="6836f-135">Notez qu’un flux déclenché manuellement n’est que l’un des nombreux types de flux.</span><span class="sxs-lookup"><span data-stu-id="6836f-135">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="6836f-136">Dans le tutoriel suivant, vous allez créer un flux qui s’exécute automatiquement lorsque vous recevez un e-mail.</span><span class="sxs-lookup"><span data-stu-id="6836f-136">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="6836f-137">Appuyez sur **Nouvelle étape**.</span><span class="sxs-lookup"><span data-stu-id="6836f-137">Press **New step**.</span></span>

6. <span data-ttu-id="6836f-138">Sélectionnez l’onglet **Standard**, puis sélectionnez **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="6836f-138">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Option dans Power Automate pour Excel Online (Business)](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="6836f-140">Sous **Actions**, sélectionnez **Exécuter le script** (Aperçu).</span><span class="sxs-lookup"><span data-stu-id="6836f-140">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Exécutez l’option d’action dans Power Automate pour exécuter le script (aperçu).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="6836f-142">Vous allez ensuite sélectionner le classeur et le script à utiliser dans l’étape de flux.</span><span class="sxs-lookup"><span data-stu-id="6836f-142">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="6836f-143">À titre de didacticiel, vous allez utiliser le classeur précédemment créé dans OneDrive, mais vous pouvez utiliser n’importe quel classeur dans un site OneDrive ou SharePoint.</span><span class="sxs-lookup"><span data-stu-id="6836f-143">For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site.</span></span> <span data-ttu-id="6836f-144">Spécifiez les paramètres suivants pour le connecteur **Exécuter le script** :</span><span class="sxs-lookup"><span data-stu-id="6836f-144">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="6836f-145">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="6836f-145">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="6836f-146">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="6836f-146">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="6836f-147">**Fichier** : MyWorkbook.xlsx *(choisi via l’Explorateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="6836f-147">**File**: MyWorkbook.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="6836f-148">**Script** : Définir la date et l’heure</span><span class="sxs-lookup"><span data-stu-id="6836f-148">**Script**: Set date and time</span></span>

    ![Les paramètres du connecteur pour exécuter un script dans Power Automate](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="6836f-150">Appuyez sur **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="6836f-150">Press **Save**.</span></span>

<span data-ttu-id="6836f-151">Votre flux est maintenant prêt à être exécuté via Power Automate.</span><span class="sxs-lookup"><span data-stu-id="6836f-151">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="6836f-152">Vous pouvez le tester à l’aide du bouton **Tester** dans l’éditeur de flux ou suivre les étapes restantes du tutoriel pour exécuter le flux à partir de votre collection de flux.</span><span class="sxs-lookup"><span data-stu-id="6836f-152">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="6836f-153">Exécuter le script via Power Automate</span><span class="sxs-lookup"><span data-stu-id="6836f-153">Run the script through Power Automate</span></span>

1. <span data-ttu-id="6836f-154">Sur la page principale de Power Automate, sélectionnez **Mes flux**.</span><span class="sxs-lookup"><span data-stu-id="6836f-154">From the main Power Automate page, select **My flows**.</span></span>

    ![Le bouton Mes flux dans Power Automate](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="6836f-156">Sélectionnez **Mon flux de tutoriel** dans la liste des flux affichée dans l’onglet **Mes flux**. Cela affiche les informations sur le flux que nous avons créé précédemment.</span><span class="sxs-lookup"><span data-stu-id="6836f-156">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="6836f-157">Appuyez sur **Exécuter**.</span><span class="sxs-lookup"><span data-stu-id="6836f-157">Press **Run**.</span></span>

    ![Le bouton Exécuter dans Power Automate](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="6836f-159">Un volet des tâches apparaîtra pour exécuter le flux.</span><span class="sxs-lookup"><span data-stu-id="6836f-159">A task pane will appear for running the flow.</span></span> <span data-ttu-id="6836f-160">Si vous êtes invité à vous **Connecter** à Excel Online, faites-le en appuyant sur **Continuer**.</span><span class="sxs-lookup"><span data-stu-id="6836f-160">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="6836f-161">Appuyez sur **Exécuter le flux**.</span><span class="sxs-lookup"><span data-stu-id="6836f-161">Press **Run flow**.</span></span> <span data-ttu-id="6836f-162">Cela exécute le flux, qui exécute le script Office associé.</span><span class="sxs-lookup"><span data-stu-id="6836f-162">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="6836f-163">Appuyez sur **Terminé**.</span><span class="sxs-lookup"><span data-stu-id="6836f-163">Press **Done**.</span></span> <span data-ttu-id="6836f-164">Vous devriez voir la section **Exécutions** s’actualiser en conséquence.</span><span class="sxs-lookup"><span data-stu-id="6836f-164">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="6836f-165">Actualisez la page pour voir les résultats de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="6836f-165">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="6836f-166">Si l’opération est réussie, accédez au classeur pour voir les cellules mises à jour.</span><span class="sxs-lookup"><span data-stu-id="6836f-166">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="6836f-167">Si l’opération a échoué, vérifiez les paramètres du flux et exécutez-le une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="6836f-167">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Production Power Automate indiquant une exécution de flux réussie](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="6836f-169">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="6836f-169">Next steps</span></span>

<span data-ttu-id="6836f-170">Suivez le tutoriel [Transférer des données aux scripts dans un flux Power Automate exécuté automatiquement](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="6836f-170">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="6836f-171">Il vous explique comment transmettre les données d’un service de flux de travail à votre script Office et comment exécuter le flux Power Automate lorsque certains événements se produisent.</span><span class="sxs-lookup"><span data-stu-id="6836f-171">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
