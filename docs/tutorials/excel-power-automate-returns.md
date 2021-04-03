---
title: Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement
description: Un didacticiel qui présente comment envoyer des e-mails de rappel en exécutant des scripts Office pour Excel sur le web via Power Automate.
ms.date: 12/15/2020
localization_priority: Priority
ms.openlocfilehash: 31ba31ddbfb36f20087be6aa7d83b1b896a698d1
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51570529"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="78ffa-103">Renvoyer les données d’un scripts vers un flux Power Automate exécuté automatiquement (aperçu)</span><span class="sxs-lookup"><span data-stu-id="78ffa-103">Return data from a script to an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="78ffa-104">Ce tutoriel vous apprend à renvoyer les informations d’un script Office pour Excel sur le web en tant qu’élément du flux de travail automatisé [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="78ffa-104">This tutorial teaches you how to return information from an Office Script for Excel on the web as part of an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="78ffa-105">Vous créerez un script qui parcoure un planning et fonctionne avec un flux pour envoyer des courriers de rappel.</span><span class="sxs-lookup"><span data-stu-id="78ffa-105">You'll make a script that looks through a schedule and works with a flow to send reminder emails.</span></span> <span data-ttu-id="78ffa-106">Ce flux s’exécutera selon un calendrier régulier, fournissant ces rappels à votre place.</span><span class="sxs-lookup"><span data-stu-id="78ffa-106">This flow will run on a regular schedule, providing these reminders on your behalf.</span></span>

> [!TIP]
> <span data-ttu-id="78ffa-107">Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="78ffa-107">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>
>
> <span data-ttu-id="78ffa-108">Si vous débutez avec Power Automate, nous vous recommandons de démarrer par les didacticiels [Appeler des scripts à partir d’un flux manuel Power Automate](excel-power-automate-manual.md) et [Transmettre des données à des scripts dans un flux automatique Power Automate (Aperçu)](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="78ffa-108">If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) and [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorials.</span></span>
>
> <span data-ttu-id="78ffa-109">[Les scripts Office utilisent TypeScript](../overview/code-editor-environment.md), et ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="78ffa-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="78ffa-110">Si vous découvrez JavaScript, nous vous conseillons de commencer par consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="78ffa-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="78ffa-111">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="78ffa-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a><span data-ttu-id="78ffa-112">Préparer le classeur</span><span class="sxs-lookup"><span data-stu-id="78ffa-112">Prepare the workbook</span></span>

1. <span data-ttu-id="78ffa-113">Téléchargez le classeur <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> dans votre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="78ffa-113">Download the workbook <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="78ffa-114">Ouvrez **on-call-rotation.xlsx** dans Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="78ffa-114">Open **on-call-rotation.xlsx** in Excel on the web.</span></span>

1. <span data-ttu-id="78ffa-115">Ajoutez une ligne au tableau avec votre nom, adresse e-mail et les dates de début et de fin qui chevauchent la date actuelle.</span><span class="sxs-lookup"><span data-stu-id="78ffa-115">Add a row to the table with your name, email address, and start and end dates that overlap with the current date.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="78ffa-116">Le script que vous écrivez utilise la première entrée correspondante dans le tableau. Vérifiez donc que votre nom figure au-dessus des lignes de la semaine actuelle.</span><span class="sxs-lookup"><span data-stu-id="78ffa-116">The script you'll write uses the first matching entry in the table, so make sure your name is above any row with the current week.</span></span>

    ![Capture d’écran d’un tableau de rotation des astreintes dans une feuille de calcul Excel](../images/power-automate-return-tutorial-1.png)

## <a name="create-an-office-script"></a><span data-ttu-id="78ffa-118">Créer un script Office</span><span class="sxs-lookup"><span data-stu-id="78ffa-118">Create an Office Script</span></span>

1. <span data-ttu-id="78ffa-119">Accédez à l’onglet **Automatiser**, puis sélectionnez **Tous les scripts**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-119">Go to the **Automate** tab and select **All Scripts**.</span></span>

1. <span data-ttu-id="78ffa-120">Sélectionnez **Nouveau script**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-120">Select **New Script**.</span></span>

1. <span data-ttu-id="78ffa-121">Nommez le script **Appeler la personne d’astreinte**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-121">Name the script **Get On-Call Person**.</span></span>

1. <span data-ttu-id="78ffa-122">Vous devez désormais avoir un script vide.</span><span class="sxs-lookup"><span data-stu-id="78ffa-122">You should now have an empty script.</span></span> <span data-ttu-id="78ffa-123">Nous utilisons le script pour obtenir l’adresse e-mail à partir de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="78ffa-123">We want to use the script to get an email address from the spreadsheet.</span></span> <span data-ttu-id="78ffa-124">Modifiez `main` pour renvoyer une chaîne, comme suit :</span><span class="sxs-lookup"><span data-stu-id="78ffa-124">Change `main` to return a string, like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. <span data-ttu-id="78ffa-125">Ensuite, nous devons obtenir toutes les données du tableau.</span><span class="sxs-lookup"><span data-stu-id="78ffa-125">Next, we need to get all the data from the table.</span></span> <span data-ttu-id="78ffa-126">Cela nous permet de parcourir chaque ligne avec le script.</span><span class="sxs-lookup"><span data-stu-id="78ffa-126">That lets us look through each row with the script.</span></span> <span data-ttu-id="78ffa-127">Ajoutez le code suivant à l’intérieur de la fonction`main`.</span><span class="sxs-lookup"><span data-stu-id="78ffa-127">Add the following code inside the `main` function.</span></span>

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. <span data-ttu-id="78ffa-128">Les dates du tableau sont stockées en utilisant le [Numéro de série de la date d’Excel](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).</span><span class="sxs-lookup"><span data-stu-id="78ffa-128">The dates in the table are stored using [Excel's date serial number](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).</span></span> <span data-ttu-id="78ffa-129">Nous convertissons ces dates en dates JavaScript pour les comparer.</span><span class="sxs-lookup"><span data-stu-id="78ffa-129">We need to convert those dates to JavaScript dates in order to compare them.</span></span> <span data-ttu-id="78ffa-130">Nous ajoutons une fonction d’assistance à notre script.</span><span class="sxs-lookup"><span data-stu-id="78ffa-130">We'll add a helper function to our script.</span></span> <span data-ttu-id="78ffa-131">Ajoutez le code suivant à l’extérieur de la fonction`main` :</span><span class="sxs-lookup"><span data-stu-id="78ffa-131">Add the following code outside of the `main` function:</span></span>

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. <span data-ttu-id="78ffa-132">Nous devons maintenant déterminer la personne d’astreinte en ce moment.</span><span class="sxs-lookup"><span data-stu-id="78ffa-132">Now, we need to figure out which person is on call right now.</span></span> <span data-ttu-id="78ffa-133">Sa ligne possède une date de début et une date de fin entourant la date actuelle.</span><span class="sxs-lookup"><span data-stu-id="78ffa-133">Their row will have a start and end date surrounding the current date.</span></span> <span data-ttu-id="78ffa-134">Nous écrivons un script pour partir du principe qu’une seule personne à la fois est d’astreinte.</span><span class="sxs-lookup"><span data-stu-id="78ffa-134">We'll write the script to assume only one person is on call at a time.</span></span> <span data-ttu-id="78ffa-135">Les scripts peuvent renvoyer des tableaux pour traiter plusieurs valeurs, mais pour l’instant, nous renvoyons la première adresse e-mail qui correspond.</span><span class="sxs-lookup"><span data-stu-id="78ffa-135">Scripts can return arrays to handle multiple values, but for now we'll return the first matching email address.</span></span> <span data-ttu-id="78ffa-136">Ajoutez la fonction suivante à la fin de la fonction `main`.</span><span class="sxs-lookup"><span data-stu-id="78ffa-136">Add the following code to the end of the `main` function.</span></span>

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. <span data-ttu-id="78ffa-137">La méthode finale doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="78ffa-137">The final script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="78ffa-138">Créer un flux de travail automatisé avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="78ffa-138">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="78ffa-139">Connectez-vous au site [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="78ffa-139">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

1. <span data-ttu-id="78ffa-140">Dans le menu qui s’affiche sur le côté gauche de l’écran, appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-140">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="78ffa-141">Cela affiche une liste des moyens de créer de nouveaux flux de travail.</span><span class="sxs-lookup"><span data-stu-id="78ffa-141">This brings you to list of ways to create new workflows.</span></span>

    ![Le bouton Créer dans Power Automate](../images/power-automate-tutorial-1.png)

1. <span data-ttu-id="78ffa-143">Sous la section **Démarrer à partir de zéro**, sélectionnez **Flux cloud planifié**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-143">Under the **Start from blank** section, select **Scheduled cloud flow**.</span></span>

    ![Le bouton Flux cloud planifié dans Power Automate](../images/power-automate-return-tutorial-2.png)

1. <span data-ttu-id="78ffa-145">Nous devons maintenant définir le planning pour ce flux.</span><span class="sxs-lookup"><span data-stu-id="78ffa-145">Now we need to set the schedule for this flow.</span></span> <span data-ttu-id="78ffa-146">Notre feuille de calcul a une nouvelle activité d’astreinte démarrant chaque lundi lors du premier semestre de 2021.</span><span class="sxs-lookup"><span data-stu-id="78ffa-146">Our spreadsheet has a new on-call assignment starting every Monday in the first half of 2021.</span></span> <span data-ttu-id="78ffa-147">Définissons le flux à exécuter en premier le lundi matin.</span><span class="sxs-lookup"><span data-stu-id="78ffa-147">Let's set the flow to run first thing Monday mornings.</span></span> <span data-ttu-id="78ffa-148">Utilisez les options suivantes pour configurer le flux à exécuter chaque semaine le lundi.</span><span class="sxs-lookup"><span data-stu-id="78ffa-148">Use the following options to configure the flow to run on Monday each week.</span></span>

    - <span data-ttu-id="78ffa-149">**Nom de flux** : Avertir la personne d’astreinte</span><span class="sxs-lookup"><span data-stu-id="78ffa-149">**Flow name**: Notify On-Call Person</span></span>
    - <span data-ttu-id="78ffa-150">**Début** : 04/01/21 à 01h00</span><span class="sxs-lookup"><span data-stu-id="78ffa-150">**Starting**: 1/4/21 at 1:00am</span></span>
    - <span data-ttu-id="78ffa-151">**Répéter tous les** : 1 semaine</span><span class="sxs-lookup"><span data-stu-id="78ffa-151">**Repeat every**: 1 Week</span></span>
    - <span data-ttu-id="78ffa-152">**Durant ces journées** : M</span><span class="sxs-lookup"><span data-stu-id="78ffa-152">**On these days**: M</span></span>

    ![Fenêtre présentant les options spécifiées pour le flux planifié](../images/power-automate-return-tutorial-3.png)

1. <span data-ttu-id="78ffa-154">Appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-154">Press **Create**.</span></span>

1. <span data-ttu-id="78ffa-155">Appuyez sur **Nouvelle étape**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-155">Press **New step**.</span></span>

1. <span data-ttu-id="78ffa-156">Sélectionnez l’onglet **Standard**, puis sélectionnez **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-156">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![L’option Excel en ligne (Entreprise) dans Power Automate](../images/power-automate-tutorial-4.png)

1. <span data-ttu-id="78ffa-158">Sous **Actions**, sélectionnez **Exécuter le script (aperçu)**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-158">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Exécuter l’option d’action de script (aperçu) dans Power Automate](../images/power-automate-tutorial-5.png)

1. <span data-ttu-id="78ffa-160">Vous allez ensuite sélectionner le classeur et le script à utiliser dans l’étape de flux.</span><span class="sxs-lookup"><span data-stu-id="78ffa-160">Next, you'll select the workbook and script to use in the flow step.</span></span> <span data-ttu-id="78ffa-161">Utilisez le classeur **rotation-des-astreintes.xlsx** que vous avez créé dans votre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="78ffa-161">Use the **on-call-rotation.xlsx** workbook you created in your OneDrive.</span></span> <span data-ttu-id="78ffa-162">Spécifiez les paramètres suivants pour le connecteur **Exécuter le script** :</span><span class="sxs-lookup"><span data-stu-id="78ffa-162">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="78ffa-163">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="78ffa-163">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="78ffa-164">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="78ffa-164">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="78ffa-165">**Fichier** : rotation-des-astreintes.xlsx *(choisi via l’Explorateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="78ffa-165">**File**: on-call-rotation.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="78ffa-166">**Script** : Obtenir la personne d’astreinte</span><span class="sxs-lookup"><span data-stu-id="78ffa-166">**Script**: Get On-Call Person</span></span>

    ![Les paramètres du connecteur pour exécuter un script dans Power Automate.](../images/power-automate-return-tutorial-4.png)

1. <span data-ttu-id="78ffa-168">Appuyez sur **Nouvelle étape**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-168">Press **New step**.</span></span>

1. <span data-ttu-id="78ffa-169">Nous allons terminer le flux en envoyant un e-mail de rappel.</span><span class="sxs-lookup"><span data-stu-id="78ffa-169">We'll end the flow by sending the reminder email.</span></span> <span data-ttu-id="78ffa-170">Sélectionnez **Envoyer un e-mail (V2)** en utilisant la barre de recherche du connecteur.</span><span class="sxs-lookup"><span data-stu-id="78ffa-170">Select **Send an email (V2)** by using the connector's search bar.</span></span> <span data-ttu-id="78ffa-171">Utilisez le contrôle **Ajouter du contenu dynamique** pour ajouter l’adresse e-mail renvoyée par le script.</span><span class="sxs-lookup"><span data-stu-id="78ffa-171">Use the **Add dynamic content** control to add the email address returned by the script.</span></span> <span data-ttu-id="78ffa-172">Cette action va étiqueter **résultat** avec l’icône Excel à côté.</span><span class="sxs-lookup"><span data-stu-id="78ffa-172">This will be labelled **result** with the Excel icon next to it.</span></span> <span data-ttu-id="78ffa-173">Vous pouvez fournir tout objet et corps de texte de votre choix.</span><span class="sxs-lookup"><span data-stu-id="78ffa-173">You can provide whatever subject and body text you'd like.</span></span>

    ![Les paramètres du connecteur pour envoyer un e-mail dans Power Automate](../images/power-automate-return-tutorial-5.png)

    > [!NOTE]
    > <span data-ttu-id="78ffa-175">Ce tutoriel utilise Outlook.</span><span class="sxs-lookup"><span data-stu-id="78ffa-175">This tutorial uses Outlook.</span></span> <span data-ttu-id="78ffa-176">N’hésitez pas à utiliser votre service de messagerie préféré, même si certaines options peuvent être différentes.</span><span class="sxs-lookup"><span data-stu-id="78ffa-176">Feel free to use your preferred email service instead, though some options may be different.</span></span>

1. <span data-ttu-id="78ffa-177">Appuyez sur **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="78ffa-177">Press **Save**.</span></span>

## <a name="test-the-script-in-power-automate"></a><span data-ttu-id="78ffa-178">Tester le script dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="78ffa-178">Test the script in Power Automate</span></span>

<span data-ttu-id="78ffa-179">Votre flux va s’exécuter chaque lundi matin.</span><span class="sxs-lookup"><span data-stu-id="78ffa-179">Your flow will run every Monday morning.</span></span> <span data-ttu-id="78ffa-180">Vous pouvez tester le script maintenant en appuyant sur le bouton **Test** dans le coin supérieur droit de l’écran.</span><span class="sxs-lookup"><span data-stu-id="78ffa-180">You can test the script now by pressing the **Test** button in the upper-right corner of the screen.</span></span> <span data-ttu-id="78ffa-181">Sélectionnez **Manuellement** et appuyez sur **Exécuter le test** pour exécuter le flux maintenant et tester le comportement.</span><span class="sxs-lookup"><span data-stu-id="78ffa-181">Select **Manually** and press **Run Test** to run the flow now and test the behavior.</span></span> <span data-ttu-id="78ffa-182">Vous devrez peut-être octroyer des autorisations à Excel et Outlook pour continuer.</span><span class="sxs-lookup"><span data-stu-id="78ffa-182">You may need to grant permissions to Excel and Outlook to continue.</span></span>

![Le bouton Test de Power Automate](../images/power-automate-return-tutorial-6.png)

> [!TIP]
> <span data-ttu-id="78ffa-184">Si votre flux ne parvient pas à envoyer un e-mail, revérifiez dans la feuille de calcul qu’une adresse e-mail valide figure dans la plage de dates actuelle en haut du tableau.</span><span class="sxs-lookup"><span data-stu-id="78ffa-184">If your flow fails to send an email, double-check in the spreadsheet that a valid email is listed for the current date range at the top of the table.</span></span>

## <a name="next-steps"></a><span data-ttu-id="78ffa-185">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="78ffa-185">Next steps</span></span>

<span data-ttu-id="78ffa-186">Visitez [Exécuter des scripts Office avec Power Automate](../develop/power-automate-integration.md) pour en savoir plus sur la connexion de scripts Office avec Power Automate.</span><span class="sxs-lookup"><span data-stu-id="78ffa-186">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="78ffa-187">Vous pouvez également consulter le [scénario type des rappels de tâches automatisés](../resources/scenarios/task-reminders.md) pour découvrir comment combiner les scripts Office et Power Automate avec les cartes adaptatives Teams.</span><span class="sxs-lookup"><span data-stu-id="78ffa-187">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
