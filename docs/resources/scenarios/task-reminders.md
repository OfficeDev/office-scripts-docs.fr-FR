---
title: 'Office Exemple de scénario de scripts : rappels de tâches automatisés'
description: Un exemple qui utilise des Power Automate et des cartes adaptatives automatise les rappels de tâches dans une feuille de calcul de gestion de projet.
ms.date: 11/30/2020
localization_priority: Normal
ms.openlocfilehash: c5515abb1e36d1bf588ab034f62dfda2625c65dc
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232857"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="c14ee-103">Office Exemple de scénario de scripts : rappels de tâches automatisés</span><span class="sxs-lookup"><span data-stu-id="c14ee-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="c14ee-104">Dans ce scénario, vous gérez un projet.</span><span class="sxs-lookup"><span data-stu-id="c14ee-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="c14ee-105">Vous utilisez une feuille de Excel pour suivre l’état de vos employés tous les mois.</span><span class="sxs-lookup"><span data-stu-id="c14ee-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="c14ee-106">Vous devez souvent rappeler aux personnes de remplir leur statut. Vous avez donc décidé d’automatiser ce processus de rappel.</span><span class="sxs-lookup"><span data-stu-id="c14ee-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="c14ee-107">Vous allez créer un flux Power Automate message aux personnes dont les champs d’état sont manquants et appliquer leurs réponses à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c14ee-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="c14ee-108">Pour ce faire, vous allez développer une paire de scripts pour gérer l’utilisation du classer.</span><span class="sxs-lookup"><span data-stu-id="c14ee-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="c14ee-109">Le premier script obtient une liste de personnes avec des états vides et le second script ajoute une chaîne d’état à la ligne de droite.</span><span class="sxs-lookup"><span data-stu-id="c14ee-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="c14ee-110">Vous utiliserez également des cartes [adaptatives Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) pour que les employés entrent leur état directement à partir de la notification.</span><span class="sxs-lookup"><span data-stu-id="c14ee-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="c14ee-111">Compétences d’écriture de scripts couvertes</span><span class="sxs-lookup"><span data-stu-id="c14ee-111">Scripting skills covered</span></span>

- <span data-ttu-id="c14ee-112">Créer des flux dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="c14ee-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="c14ee-113">Transmettre des données à des scripts</span><span class="sxs-lookup"><span data-stu-id="c14ee-113">Pass data to scripts</span></span>
- <span data-ttu-id="c14ee-114">Renvoyer des données à partir de scripts</span><span class="sxs-lookup"><span data-stu-id="c14ee-114">Return data from scripts</span></span>
- <span data-ttu-id="c14ee-115">Teams Cartes adaptatives</span><span class="sxs-lookup"><span data-stu-id="c14ee-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="c14ee-116">Tables</span><span class="sxs-lookup"><span data-stu-id="c14ee-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c14ee-117">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c14ee-117">Prerequisites</span></span>

<span data-ttu-id="c14ee-118">Ce scénario utilise [Power Automate](https://flow.microsoft.com) et [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="c14ee-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="c14ee-119">Vous aurez besoin des deux associés au compte que vous utilisez pour le développement de Office scripts.</span><span class="sxs-lookup"><span data-stu-id="c14ee-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="c14ee-120">Pour obtenir un accès gratuit à un abonnement Microsoft Développeur pour en savoir plus sur ces applications et travailler avec celles-ci, envisagez de rejoindre le programme [Microsoft 365 développeur microsoft.](https://developer.microsoft.com/microsoft-365/dev-program)</span><span class="sxs-lookup"><span data-stu-id="c14ee-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="c14ee-121">Instructions d’installation</span><span class="sxs-lookup"><span data-stu-id="c14ee-121">Setup instructions</span></span>

1. <span data-ttu-id="c14ee-122">Téléchargez <a href="task-reminders.xlsx">task-reminders.xlsx</a> sur votre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="c14ee-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="c14ee-123">Ouvrez le Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="c14ee-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="c14ee-124">Sous **l’onglet Automatiser,** ouvrez **Tous les scripts.**</span><span class="sxs-lookup"><span data-stu-id="c14ee-124">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="c14ee-125">Tout d’abord, nous avons besoin d’un script pour obtenir tous les employés dont les rapports d’état sont manquants dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c14ee-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="c14ee-126">Dans le **volet Des tâches de** l’Éditeur de code, appuyez sur Nouveau **script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="c14ee-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="c14ee-127">Enregistrez le script avec le nom **Get People**.</span><span class="sxs-lookup"><span data-stu-id="c14ee-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="c14ee-128">Ensuite, nous avons besoin d’un second script pour traiter les cartes de rapport d’état et placer les nouvelles informations dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c14ee-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="c14ee-129">Dans le **volet Des tâches de** l’Éditeur de code, appuyez sur Nouveau **script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="c14ee-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

7. <span data-ttu-id="c14ee-130">Enregistrez le script sous le nom **Enregistrer l’état**.</span><span class="sxs-lookup"><span data-stu-id="c14ee-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="c14ee-131">Maintenant, nous devons créer le flux.</span><span class="sxs-lookup"><span data-stu-id="c14ee-131">Now, we need to create the flow.</span></span> <span data-ttu-id="c14ee-132">Ouvrez [Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="c14ee-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="c14ee-133">Si vous n’avez pas encore créé de flux, consultez notre didacticiel Commencez à utiliser des [scripts](../../tutorials/excel-power-automate-manual.md) Power Automate pour en savoir plus sur les bases.</span><span class="sxs-lookup"><span data-stu-id="c14ee-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="c14ee-134">Créez un **flux instantané.**</span><span class="sxs-lookup"><span data-stu-id="c14ee-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="c14ee-135">Choose **Manually trigger a flow** from the options and press **Create**.</span><span class="sxs-lookup"><span data-stu-id="c14ee-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="c14ee-136">Le flux doit appeler le script **Obtenir des** personnes pour obtenir tous les employés avec des champs d’état vides.</span><span class="sxs-lookup"><span data-stu-id="c14ee-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="c14ee-137">Appuyez **sur Nouvelle étape** et **sélectionnez Excel Online (Entreprise).**</span><span class="sxs-lookup"><span data-stu-id="c14ee-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="c14ee-138">Sous **Actions**, sélectionnez **Exécuter le script (aperçu)**.</span><span class="sxs-lookup"><span data-stu-id="c14ee-138">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="c14ee-139">Fournissez les entrées suivantes pour l’étape de flux :</span><span class="sxs-lookup"><span data-stu-id="c14ee-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="c14ee-140">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="c14ee-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="c14ee-141">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="c14ee-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="c14ee-142">**Fichier**: task-reminders.xlsx *(choisi via le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="c14ee-142">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="c14ee-143">**Script**: obtenir des personnes</span><span class="sxs-lookup"><span data-stu-id="c14ee-143">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Flux de Power Automate montrant la première étape de flux de script d’exécuter":::

12. <span data-ttu-id="c14ee-145">Ensuite, le flux doit traiter chaque employé dans le tableau renvoyé par le script.</span><span class="sxs-lookup"><span data-stu-id="c14ee-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="c14ee-146">Appuyez **sur Nouvelle étape** et sélectionnez Publier une carte **adaptative à un utilisateur Teams et attendez une réponse.**</span><span class="sxs-lookup"><span data-stu-id="c14ee-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="c14ee-147">Pour le **champ Destinataire,** ajoutez **le courrier électronique** à partir du contenu dynamique (la sélection Excel logo).</span><span class="sxs-lookup"><span data-stu-id="c14ee-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="c14ee-148">**L’ajout d’un** courrier électronique entraîne le fait que l’étape du flux soit entourée d’une **application à chaque** bloc.</span><span class="sxs-lookup"><span data-stu-id="c14ee-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="c14ee-149">Cela signifie que le tableau est itéré par Power Automate.</span><span class="sxs-lookup"><span data-stu-id="c14ee-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="c14ee-150">L’envoi d’une carte adaptative nécessite que le JSON de la carte soit fourni en tant que **message.**</span><span class="sxs-lookup"><span data-stu-id="c14ee-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="c14ee-151">Vous pouvez utiliser le Concepteur de [cartes adaptatives pour](https://adaptivecards.io/designer/) créer des cartes personnalisées.</span><span class="sxs-lookup"><span data-stu-id="c14ee-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="c14ee-152">Pour cet exemple, utilisez le JSON suivant.</span><span class="sxs-lookup"><span data-stu-id="c14ee-152">For this sample, use the following JSON.</span></span>  

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

15. <span data-ttu-id="c14ee-153">Remplissez les champs restants comme suit :</span><span class="sxs-lookup"><span data-stu-id="c14ee-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="c14ee-154">**Message de mise à** jour : merci d’avoir envoyé votre rapport d’état.</span><span class="sxs-lookup"><span data-stu-id="c14ee-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="c14ee-155">Votre réponse a été ajoutée avec succès à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c14ee-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="c14ee-156">**Doit mettre à jour la carte**: Oui</span><span class="sxs-lookup"><span data-stu-id="c14ee-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="c14ee-157">Dans le **bloc Appliquer à chaque** bloc, après avoir publié une carte adaptative à un utilisateur **Teams** et attendre une réponse, appuyez sur Ajouter **une action.**</span><span class="sxs-lookup"><span data-stu-id="c14ee-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="c14ee-158">Sélectionnez **Excel Online (Entreprise).**</span><span class="sxs-lookup"><span data-stu-id="c14ee-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="c14ee-159">Sous **Actions**, sélectionnez **Exécuter le script (aperçu)**.</span><span class="sxs-lookup"><span data-stu-id="c14ee-159">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="c14ee-160">Fournissez les entrées suivantes pour l’étape de flux :</span><span class="sxs-lookup"><span data-stu-id="c14ee-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="c14ee-161">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="c14ee-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="c14ee-162">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="c14ee-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="c14ee-163">**Fichier**: task-reminders.xlsx *(choisi via le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="c14ee-163">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="c14ee-164">**Script**: Enregistrer l’état</span><span class="sxs-lookup"><span data-stu-id="c14ee-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="c14ee-165">**senderEmail**: e-mail *(contenu dynamique de Excel)*</span><span class="sxs-lookup"><span data-stu-id="c14ee-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="c14ee-166">**statusReportResponse**: réponse *(contenu dynamique de Teams)*</span><span class="sxs-lookup"><span data-stu-id="c14ee-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Flux Power Automate montrant l’application à chaque étape":::

17. <span data-ttu-id="c14ee-168">Enregistrez le flux.</span><span class="sxs-lookup"><span data-stu-id="c14ee-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="c14ee-169">Exécution du flux</span><span class="sxs-lookup"><span data-stu-id="c14ee-169">Running the flow</span></span>

<span data-ttu-id="c14ee-170">Pour tester le flux, assurez-vous que les lignes de tableau dont l’état est vide utilisent une adresse de messagerie liée à un compte Teams (vous devez probablement utiliser votre propre adresse e-mail lors du test).</span><span class="sxs-lookup"><span data-stu-id="c14ee-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="c14ee-171">Vous pouvez sélectionner **Test** à partir du concepteur de flux ou exécuter le flux à partir de la page **Mes flux.**</span><span class="sxs-lookup"><span data-stu-id="c14ee-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="c14ee-172">Après avoir commencé le flux et accepté l’utilisation des connexions requises, vous devez recevoir une carte adaptative de Power Automate à Teams.</span><span class="sxs-lookup"><span data-stu-id="c14ee-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="c14ee-173">Une fois que vous avez rempli le champ d’état dans la carte, le flux continue et met à jour la feuille de calcul avec l’état que vous fournissez.</span><span class="sxs-lookup"><span data-stu-id="c14ee-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="c14ee-174">Avant d’exécution du flux</span><span class="sxs-lookup"><span data-stu-id="c14ee-174">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Feuille de calcul avec un rapport d’état contenant une entrée d’état manquante":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="c14ee-176">Réception de la carte adaptative</span><span class="sxs-lookup"><span data-stu-id="c14ee-176">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Une carte adaptative dans Teams demande de mise à jour de l’état à l’employé":::

### <a name="after-running-the-flow"></a><span data-ttu-id="c14ee-178">Après l’exécution du flux</span><span class="sxs-lookup"><span data-stu-id="c14ee-178">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Feuille de calcul avec un rapport d’état avec une entrée d’état maintenant remplie":::
