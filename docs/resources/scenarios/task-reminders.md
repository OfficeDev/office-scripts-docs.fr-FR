---
title: 'Office Scénario d’exemple de scripts : rappels de tâches automatisés'
description: Un échantillon qui utilise des cartes Power Automate adaptatives automatise les rappels de tâches dans une feuille de calcul de gestion de projet.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c254a627da8442c0974263908a41275182740b6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545601"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="e22d5-103">Office Scénario d’exemple de scripts : rappels de tâches automatisés</span><span class="sxs-lookup"><span data-stu-id="e22d5-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="e22d5-104">Dans ce scénario, vous gérez un projet.</span><span class="sxs-lookup"><span data-stu-id="e22d5-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="e22d5-105">Vous utilisez une feuille Excel pour suivre l’état de vos employés chaque mois.</span><span class="sxs-lookup"><span data-stu-id="e22d5-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="e22d5-106">Vous devez souvent rappeler aux gens de remplir leur statut, vous avez donc décidé d’automatiser ce processus de rappel.</span><span class="sxs-lookup"><span data-stu-id="e22d5-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="e22d5-107">Vous créerez un flux de Power Automate pour envoyer des messages aux personnes ayant des champs d’état manquants et appliquerez leurs réponses à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e22d5-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="e22d5-108">Pour ce faire, vous développerez une paire de scripts pour gérer le travail avec le cahier de travail.</span><span class="sxs-lookup"><span data-stu-id="e22d5-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="e22d5-109">Le premier script reçoit une liste de personnes ayant des statuts vierges et le deuxième script ajoute une chaîne de statut à la bonne ligne.</span><span class="sxs-lookup"><span data-stu-id="e22d5-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="e22d5-110">Vous utiliserez également les cartes [adaptatives Teams pour que](/microsoftteams/platform/task-modules-and-cards/what-are-cards) les employés saisiront leur statut directement à partir de la notification.</span><span class="sxs-lookup"><span data-stu-id="e22d5-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="e22d5-111">Compétences de script couvertes</span><span class="sxs-lookup"><span data-stu-id="e22d5-111">Scripting skills covered</span></span>

- <span data-ttu-id="e22d5-112">Créer des flux dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="e22d5-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="e22d5-113">Transmettre des données aux scripts</span><span class="sxs-lookup"><span data-stu-id="e22d5-113">Pass data to scripts</span></span>
- <span data-ttu-id="e22d5-114">Renvoyer les données des scripts</span><span class="sxs-lookup"><span data-stu-id="e22d5-114">Return data from scripts</span></span>
- <span data-ttu-id="e22d5-115">Teams Cartes adaptatives</span><span class="sxs-lookup"><span data-stu-id="e22d5-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="e22d5-116">Tables</span><span class="sxs-lookup"><span data-stu-id="e22d5-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e22d5-117">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e22d5-117">Prerequisites</span></span>

<span data-ttu-id="e22d5-118">Ce scénario utilise [Power Automate](https://flow.microsoft.com) et [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="e22d5-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="e22d5-119">Vous aurez besoin à la fois associé au compte que vous utilisez pour développer Office scripts.</span><span class="sxs-lookup"><span data-stu-id="e22d5-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="e22d5-120">Pour accéder gratuitement à un abonnement Microsoft Developer pour en savoir plus sur ces applications et y travailler, envisagez de rejoindre [le Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span><span class="sxs-lookup"><span data-stu-id="e22d5-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="e22d5-121">Instructions d’installation</span><span class="sxs-lookup"><span data-stu-id="e22d5-121">Setup instructions</span></span>

1. <span data-ttu-id="e22d5-122">Téléchargez <a href="task-reminders.xlsx">task-reminders.xlsx</a> sur votre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="e22d5-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="e22d5-123">Ouvrez le cahier de travail en Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="e22d5-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="e22d5-124">Sous **l’onglet Automate,** ouvrez **tous les scripts**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="e22d5-125">Tout d’abord, nous avons besoin d’un script pour obtenir tous les employés avec des rapports d’état qui manquent à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e22d5-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="e22d5-126">Dans le volet de tâche de l’éditeur de **code,** **appuyez sur Nouveau Script** et coller le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="e22d5-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

5. <span data-ttu-id="e22d5-127">Enregistrer le script avec le nom **Get People**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="e22d5-128">Ensuite, nous avons besoin d’un deuxième script pour traiter les bulletins d’état et mettre les nouvelles informations dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e22d5-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="e22d5-129">Dans le volet de tâche de l’éditeur de **code,** **appuyez sur Nouveau Script** et coller le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="e22d5-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

7. <span data-ttu-id="e22d5-130">Enregistrez le script avec le nom **Enregistrer le statut**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="e22d5-131">Maintenant, nous devons créer le flux.</span><span class="sxs-lookup"><span data-stu-id="e22d5-131">Now, we need to create the flow.</span></span> <span data-ttu-id="e22d5-132">Ouvrez [Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="e22d5-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="e22d5-133">Si vous n’avez pas créé un flux avant, s’il vous plaît consulter notre [tutoriel Commencez à utiliser des scripts avec Power Automate](../../tutorials/excel-power-automate-manual.md) pour apprendre les bases.</span><span class="sxs-lookup"><span data-stu-id="e22d5-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="e22d5-134">Créez un nouveau **flux instantané**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="e22d5-135">Choisissez **de déclencher manuellement un flux à partir** des options et appuyez sur **Créer**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="e22d5-136">Le flux doit appeler le script **Get People pour** obtenir tous les employés avec des champs de statut vides.</span><span class="sxs-lookup"><span data-stu-id="e22d5-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="e22d5-137">Appuyez **sur Nouvelle** étape et **sélectionnez Excel en ligne (Affaires)**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="e22d5-138">Sous **Actions**, sélectionnez **Script d’exécuter**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-138">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="e22d5-139">Fournissez les entrées suivantes pour l’étape de flux :</span><span class="sxs-lookup"><span data-stu-id="e22d5-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="e22d5-140">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="e22d5-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="e22d5-141">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="e22d5-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="e22d5-142">**Fichier**: task-reminders.xlsx *(Choisi par le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="e22d5-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="e22d5-143">**Script**: Get People</span><span class="sxs-lookup"><span data-stu-id="e22d5-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Le flux Power Automate affichage de la première étape de flux de script s’exécutent":::

12. <span data-ttu-id="e22d5-145">Ensuite, le flux doit traiter chaque employé dans le tableau retourné par le script.</span><span class="sxs-lookup"><span data-stu-id="e22d5-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="e22d5-146">Appuyez **sur Nouvelle** étape **et sélectionnez Poster une carte adaptative à Teams utilisateur et attendre une réponse**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="e22d5-147">Pour le **champ** Destinataire, ajoutez **l’e-mail** du contenu dynamique (la sélection aura le logo Excel par elle).</span><span class="sxs-lookup"><span data-stu-id="e22d5-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="e22d5-148">**L’ajout** d’e-mail provoque l’étape de flux d’être **entouré par une application à** chaque bloc.</span><span class="sxs-lookup"><span data-stu-id="e22d5-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="e22d5-149">Cela signifie que le tableau sera itéré par Power Automate.</span><span class="sxs-lookup"><span data-stu-id="e22d5-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="e22d5-150">L’envoi d’une carte adaptative exige que le JSON de la carte soit fourni sous forme de **message.**</span><span class="sxs-lookup"><span data-stu-id="e22d5-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="e22d5-151">Vous pouvez utiliser le concepteur [de cartes adaptatives](https://adaptivecards.io/designer/) pour créer des cartes personnalisées.</span><span class="sxs-lookup"><span data-stu-id="e22d5-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="e22d5-152">Pour cet échantillon, utilisez le JSON suivant.</span><span class="sxs-lookup"><span data-stu-id="e22d5-152">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

15. <span data-ttu-id="e22d5-153">Remplissez les champs restants comme suit :</span><span class="sxs-lookup"><span data-stu-id="e22d5-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="e22d5-154">**Message de mise** à jour : Merci d’avoir soumis votre rapport d’état.</span><span class="sxs-lookup"><span data-stu-id="e22d5-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="e22d5-155">Votre réponse a été ajoutée avec succès à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e22d5-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="e22d5-156">**Devrait mettre à jour la** carte : Oui</span><span class="sxs-lookup"><span data-stu-id="e22d5-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="e22d5-157">Dans **l’apply à chaque** bloc, en suivant **la publication d’une carte adaptative à un utilisateur Teams et attendre une réponse, appuyez** **sur Ajouter une action**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="e22d5-158">Sélectionnez **Excel en ligne (Affaires)**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="e22d5-159">Sous **Actions**, sélectionnez **Script d’exécuter**.</span><span class="sxs-lookup"><span data-stu-id="e22d5-159">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="e22d5-160">Fournissez les entrées suivantes pour l’étape de flux :</span><span class="sxs-lookup"><span data-stu-id="e22d5-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="e22d5-161">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="e22d5-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="e22d5-162">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="e22d5-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="e22d5-163">**Fichier**: task-reminders.xlsx *(Choisi par le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="e22d5-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="e22d5-164">**Script**: Enregistrer le statut</span><span class="sxs-lookup"><span data-stu-id="e22d5-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="e22d5-165">**senderEmail**: e-mail *(contenu dynamique de Excel)*</span><span class="sxs-lookup"><span data-stu-id="e22d5-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="e22d5-166">**statusReportResponse**: réponse *(contenu dynamique de Teams)*</span><span class="sxs-lookup"><span data-stu-id="e22d5-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Le Power Automate de la circulation montrant l’application à chaque étape":::

17. <span data-ttu-id="e22d5-168">Enregistrez le flux.</span><span class="sxs-lookup"><span data-stu-id="e22d5-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="e22d5-169">Exécution du flux</span><span class="sxs-lookup"><span data-stu-id="e22d5-169">Running the flow</span></span>

<span data-ttu-id="e22d5-170">Pour tester le flux, assurez-vous que toutes les lignes de table avec l’état vierge utilisent une adresse e-mail liée à un compte Teams (vous devriez probablement utiliser votre propre adresse e-mail lors des tests).</span><span class="sxs-lookup"><span data-stu-id="e22d5-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="e22d5-171">Vous pouvez sélectionner Test **à partir** du concepteur de flux, ou exécuter le flux à partir de la page **Mes flux.**</span><span class="sxs-lookup"><span data-stu-id="e22d5-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="e22d5-172">Après avoir commencé le flux et accepté l’utilisation des connexions requises, vous devez recevoir une carte adaptative de Power Automate à Teams.</span><span class="sxs-lookup"><span data-stu-id="e22d5-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="e22d5-173">Une fois que vous remplissez le champ d’état de la carte, le flux se poursuivra et mettra à jour la feuille de calcul avec l’état que vous fournissez.</span><span class="sxs-lookup"><span data-stu-id="e22d5-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="e22d5-174">Avant d’exécuter le flux</span><span class="sxs-lookup"><span data-stu-id="e22d5-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Une feuille de travail avec un rapport d’état contenant une entrée d’état manquante":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="e22d5-176">Réception de la carte adaptative</span><span class="sxs-lookup"><span data-stu-id="e22d5-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Une carte adaptative en Teams à l’employé pour une mise à jour de statut":::

### <a name="after-running-the-flow"></a><span data-ttu-id="e22d5-178">Après avoir fait fonctionner le flux</span><span class="sxs-lookup"><span data-stu-id="e22d5-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Une feuille de travail avec un rapport d’état avec une entrée de statut maintenant remplie":::
