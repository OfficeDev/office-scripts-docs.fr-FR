---
title: 'Office Exemple de scénario de scripts : rappels de tâches automatisés'
description: Un exemple qui utilise des Power Automate et des cartes adaptatives automatise les rappels de tâches dans une feuille de calcul de gestion de projet.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: cf25b81ad44bbe963083f6a8346c0fd59a514305
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313980"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="b425f-103">Office Exemple de scénario de scripts : rappels de tâches automatisés</span><span class="sxs-lookup"><span data-stu-id="b425f-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="b425f-104">Dans ce scénario, vous gérez un projet.</span><span class="sxs-lookup"><span data-stu-id="b425f-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="b425f-105">Vous utilisez une feuille Excel pour suivre l’état de vos employés tous les mois.</span><span class="sxs-lookup"><span data-stu-id="b425f-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="b425f-106">Vous devez souvent rappeler aux personnes de remplir leur statut. Vous avez donc décidé d’automatiser ce processus de rappel.</span><span class="sxs-lookup"><span data-stu-id="b425f-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="b425f-107">Vous allez créer un flux Power Automate message aux personnes dont les champs d’état sont manquants et appliquer leurs réponses à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b425f-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="b425f-108">Pour ce faire, vous allez développer une paire de scripts pour gérer l’utilisation du classer.</span><span class="sxs-lookup"><span data-stu-id="b425f-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="b425f-109">Le premier script obtient la liste des personnes dont l’état est vide et le second ajoute une chaîne d’état à la ligne de droite.</span><span class="sxs-lookup"><span data-stu-id="b425f-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="b425f-110">Vous utiliserez également des cartes [adaptatives Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) pour que les employés entrent leur état directement à partir de la notification.</span><span class="sxs-lookup"><span data-stu-id="b425f-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="b425f-111">Compétences d’écriture de scripts couvertes</span><span class="sxs-lookup"><span data-stu-id="b425f-111">Scripting skills covered</span></span>

- <span data-ttu-id="b425f-112">Créer des flux dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="b425f-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="b425f-113">Transmettre des données à des scripts</span><span class="sxs-lookup"><span data-stu-id="b425f-113">Pass data to scripts</span></span>
- <span data-ttu-id="b425f-114">Renvoyer des données à partir de scripts</span><span class="sxs-lookup"><span data-stu-id="b425f-114">Return data from scripts</span></span>
- <span data-ttu-id="b425f-115">Teams Cartes adaptatives</span><span class="sxs-lookup"><span data-stu-id="b425f-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="b425f-116">Tables</span><span class="sxs-lookup"><span data-stu-id="b425f-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b425f-117">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="b425f-117">Prerequisites</span></span>

<span data-ttu-id="b425f-118">Ce scénario utilise [Power Automate](https://flow.microsoft.com) et [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="b425f-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="b425f-119">Vous aurez besoin des deux associés au compte que vous utilisez pour développer Office Scripts.</span><span class="sxs-lookup"><span data-stu-id="b425f-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="b425f-120">Pour un accès gratuit à un abonnement Microsoft Développeur pour en savoir plus sur ces applications et travailler avec celles-ci, envisagez de rejoindre le programme [Microsoft 365 développeur microsoft.](https://developer.microsoft.com/microsoft-365/dev-program)</span><span class="sxs-lookup"><span data-stu-id="b425f-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="b425f-121">Instructions d’installation</span><span class="sxs-lookup"><span data-stu-id="b425f-121">Setup instructions</span></span>

1. <span data-ttu-id="b425f-122">Téléchargez <a href="task-reminders.xlsx">task-reminders.xlsx</a> sur votre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="b425f-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

1. <span data-ttu-id="b425f-123">Ouvrez le Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="b425f-123">Open the workbook in Excel on the web.</span></span>

1. <span data-ttu-id="b425f-124">Tout d’abord, nous avons besoin d’un script pour obtenir tous les employés dont les rapports d’état sont manquants dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b425f-124">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="b425f-125">Sous **l’onglet Automatiser,** sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="b425f-125">Under the **Automate** tab, select **New Script** and paste the following script into the editor.</span></span>

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

1. <span data-ttu-id="b425f-126">Enregistrez le script avec le nom **Get People**.</span><span class="sxs-lookup"><span data-stu-id="b425f-126">Save the script with the name **Get People**.</span></span>

1. <span data-ttu-id="b425f-127">Ensuite, nous avons besoin d’un second script pour traiter les cartes de rapport d’état et placer les nouvelles informations dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b425f-127">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="b425f-128">Dans le volet Des tâches de l’Éditeur de code, sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="b425f-128">In the Code Editor task pane, select **New Script** and paste the following script into the editor.</span></span>

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

1. <span data-ttu-id="b425f-129">Enregistrez le script sous le nom **Enregistrer l’état**.</span><span class="sxs-lookup"><span data-stu-id="b425f-129">Save the script with the name **Save Status**.</span></span>

1. <span data-ttu-id="b425f-130">Maintenant, nous devons créer le flux.</span><span class="sxs-lookup"><span data-stu-id="b425f-130">Now, we need to create the flow.</span></span> <span data-ttu-id="b425f-131">Ouvrez [Power Automate](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="b425f-131">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="b425f-132">Si vous n’avez pas encore créé de flux, consultez notre didacticiel Commencez à utiliser des [scripts](../../tutorials/excel-power-automate-manual.md) Power Automate pour en savoir plus sur les bases.</span><span class="sxs-lookup"><span data-stu-id="b425f-132">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

1. <span data-ttu-id="b425f-133">Créez un **flux instantané.**</span><span class="sxs-lookup"><span data-stu-id="b425f-133">Create a new **Instant flow**.</span></span>

1. <span data-ttu-id="b425f-134">Sélectionnez **Déclencher manuellement un flux à** partir des options et sélectionnez **Créer.**</span><span class="sxs-lookup"><span data-stu-id="b425f-134">Choose **Manually trigger a flow** from the options and select **Create**.</span></span>

1. <span data-ttu-id="b425f-135">Le flux doit appeler le script **Obtenir des** personnes pour obtenir tous les employés avec des champs d’état vides.</span><span class="sxs-lookup"><span data-stu-id="b425f-135">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="b425f-136">Sélectionnez **Nouvelle étape,** puis **sélectionnez Excel Online (Entreprise).**</span><span class="sxs-lookup"><span data-stu-id="b425f-136">Select **New step**, then select **Excel Online (Business)**.</span></span> <span data-ttu-id="b425f-137">Sous **Actions**, sélectionnez **Exécuter le script**.</span><span class="sxs-lookup"><span data-stu-id="b425f-137">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="b425f-138">Fournissez les entrées suivantes pour l’étape de flux :</span><span class="sxs-lookup"><span data-stu-id="b425f-138">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="b425f-139">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="b425f-139">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="b425f-140">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="b425f-140">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="b425f-141">**Fichier**: task-reminders.xlsx *(choisi via le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="b425f-141">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="b425f-142">**Script**: obtenir des personnes</span><span class="sxs-lookup"><span data-stu-id="b425f-142">**Script**: Get People</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Flux Power Automate montrant la première étape du flux de script d’exécuter.":::

1. <span data-ttu-id="b425f-144">Ensuite, le flux doit traiter chaque employé dans le tableau renvoyé par le script.</span><span class="sxs-lookup"><span data-stu-id="b425f-144">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="b425f-145">Sélectionnez **Nouvelle étape,** puis choisissez Publier une carte adaptative à **un utilisateur Teams et attendre une réponse.**</span><span class="sxs-lookup"><span data-stu-id="b425f-145">Select **New step**, then choose **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

1. <span data-ttu-id="b425f-146">Pour le **champ Destinataire,** ajoutez **le courrier** électronique à partir du contenu dynamique (la sélection Excel logo).</span><span class="sxs-lookup"><span data-stu-id="b425f-146">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="b425f-147">**L’ajout d’un** message électronique entraîne le fait que l’étape du flux soit entourée d’une **application à chaque** bloc.</span><span class="sxs-lookup"><span data-stu-id="b425f-147">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="b425f-148">Cela signifie que le tableau sera itéré par Power Automate.</span><span class="sxs-lookup"><span data-stu-id="b425f-148">That means the array will be iterated over by Power Automate.</span></span>

1. <span data-ttu-id="b425f-149">L’envoi d’une carte adaptative nécessite que le JSON de la carte soit fourni en tant que **message.**</span><span class="sxs-lookup"><span data-stu-id="b425f-149">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="b425f-150">Vous pouvez utiliser le Concepteur de [cartes adaptatives pour](https://adaptivecards.io/designer/) créer des cartes personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b425f-150">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="b425f-151">Pour cet exemple, utilisez le JSON suivant.</span><span class="sxs-lookup"><span data-stu-id="b425f-151">For this sample, use the following JSON.</span></span>  

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

1. <span data-ttu-id="b425f-152">Remplissez les champs restants comme suit :</span><span class="sxs-lookup"><span data-stu-id="b425f-152">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="b425f-153">**Message de mise à** jour : merci d’avoir envoyé votre rapport d’état.</span><span class="sxs-lookup"><span data-stu-id="b425f-153">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="b425f-154">Votre réponse a été ajoutée avec succès à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b425f-154">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="b425f-155">**Doit mettre à jour la carte**: Oui</span><span class="sxs-lookup"><span data-stu-id="b425f-155">**Should update card**: Yes</span></span>

1. <span data-ttu-id="b425f-156">In the **Apply to each** block, following the Post an Adaptive Card to a Teams user and wait for a **response**, select Add **an action**.</span><span class="sxs-lookup"><span data-stu-id="b425f-156">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, select **Add an action**.</span></span> <span data-ttu-id="b425f-157">Sélectionnez **Excel Online (Entreprise).**</span><span class="sxs-lookup"><span data-stu-id="b425f-157">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="b425f-158">Sous **Actions**, sélectionnez **Exécuter le script**.</span><span class="sxs-lookup"><span data-stu-id="b425f-158">Under **Actions**, select **Run script**.</span></span> <span data-ttu-id="b425f-159">Fournissez les entrées suivantes pour l’étape de flux :</span><span class="sxs-lookup"><span data-stu-id="b425f-159">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="b425f-160">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="b425f-160">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="b425f-161">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="b425f-161">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="b425f-162">**Fichier**: task-reminders.xlsx *(choisi via le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="b425f-162">**File**: task-reminders.xlsx *(Chosen through the file browser)*</span></span>
    - <span data-ttu-id="b425f-163">**Script**: enregistrer l’état</span><span class="sxs-lookup"><span data-stu-id="b425f-163">**Script**: Save Status</span></span>
    - <span data-ttu-id="b425f-164">**senderEmail**: courrier *électronique (contenu dynamique de Excel)*</span><span class="sxs-lookup"><span data-stu-id="b425f-164">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="b425f-165">**statusReportResponse**: réponse *(contenu dynamique de Teams)*</span><span class="sxs-lookup"><span data-stu-id="b425f-165">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Le Power Automate flux montrant l’application à chaque étape.":::

1. <span data-ttu-id="b425f-167">Enregistrez le flux.</span><span class="sxs-lookup"><span data-stu-id="b425f-167">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="b425f-168">Exécution du flux</span><span class="sxs-lookup"><span data-stu-id="b425f-168">Running the flow</span></span>

<span data-ttu-id="b425f-169">Pour tester le flux, assurez-vous que les lignes de tableau dont l’état est vide utilisent une adresse de messagerie liée à un compte Teams (vous devez probablement utiliser votre propre adresse e-mail lors du test).</span><span class="sxs-lookup"><span data-stu-id="b425f-169">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span> <span data-ttu-id="b425f-170">Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.</span><span class="sxs-lookup"><span data-stu-id="b425f-170">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>

<span data-ttu-id="b425f-171">Vous devez recevoir une carte adaptative de Power Automate à Teams.</span><span class="sxs-lookup"><span data-stu-id="b425f-171">You should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="b425f-172">Une fois que vous avez rempli le champ d’état dans la carte, le flux continue et met à jour la feuille de calcul avec l’état que vous fournissez.</span><span class="sxs-lookup"><span data-stu-id="b425f-172">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="b425f-173">Avant d’exécution du flux</span><span class="sxs-lookup"><span data-stu-id="b425f-173">Before running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Feuille de calcul avec un rapport d’état contenant une entrée d’état manquante.":::

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="b425f-175">Réception de la carte adaptative</span><span class="sxs-lookup"><span data-stu-id="b425f-175">Receiving the Adaptive Card</span></span>

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Une carte adaptative Teams demande à l’employé une mise à jour de l’état.":::

### <a name="after-running-the-flow"></a><span data-ttu-id="b425f-177">Après l’exécution du flux</span><span class="sxs-lookup"><span data-stu-id="b425f-177">After running the flow</span></span>

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Feuille de calcul avec un rapport d’état avec une entrée d’état maintenant remplie.":::
