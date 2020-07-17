---
title: Exécuter automatiquement des scripts avec automate d’alimentation automatisée des flux
description: Didacticiel sur l’exécution de scripts Office pour Excel sur le Web via automate d’alimentation à l’aide d’un déclencheur externe automatique (réception de courriers électroniques via Outlook).
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: fc98fb36fd5a8c5ef10bc3b767d6f5add0306246
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081627"
---
# <a name="automatically-run-scripts-with-automated-power-automate-flows-preview"></a><span data-ttu-id="e2713-103">Exécuter automatiquement des scripts avec automate d’alimentation automatique flux (aperçu)</span><span class="sxs-lookup"><span data-stu-id="e2713-103">Automatically run scripts with automated Power Automate flows (preview)</span></span>

<span data-ttu-id="e2713-104">Ce didacticiel vous apprend à utiliser un script Office pour Excel sur le Web avec un flux de travail Automated [Power](https://flow.microsoft.com) .</span><span class="sxs-lookup"><span data-stu-id="e2713-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="e2713-105">Votre script s’exécute automatiquement chaque fois que vous recevez un courrier électronique, en enregistrant des informations à partir du courrier électronique dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="e2713-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e2713-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="e2713-106">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="e2713-107">Ce didacticiel suppose que vous avez terminé l' [exécution des scripts Office dans Excel sur le Web avec le didacticiel Power Automated](excel-power-automate-manual.md) .</span><span class="sxs-lookup"><span data-stu-id="e2713-107">This tutorial assumes you have completed the [Run Office Scripts in Excel on the web with Power Automate](excel-power-automate-manual.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="e2713-108">Préparation du classeur</span><span class="sxs-lookup"><span data-stu-id="e2713-108">Prepare the workbook</span></span>

<span data-ttu-id="e2713-109">Power automate ne peut pas utiliser de [références relatives](../develop/power-automate-integration.md#avoid-using-relative-references) comme `Workbook.getActiveWorksheet` pour accéder aux composants du classeur.</span><span class="sxs-lookup"><span data-stu-id="e2713-109">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="e2713-110">Par conséquent, nous avons besoin d’un classeur et d’une feuille de calcul avec des noms cohérents pour automate de puissance à référencer.</span><span class="sxs-lookup"><span data-stu-id="e2713-110">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="e2713-111">Créez un classeur nommé **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="e2713-111">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="e2713-112">Accédez à l’onglet **automatiser** et sélectionnez **éditeur de code**.</span><span class="sxs-lookup"><span data-stu-id="e2713-112">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="e2713-113">Sélectionnez **nouveau script**.</span><span class="sxs-lookup"><span data-stu-id="e2713-113">Select **New Script**.</span></span>

4. <span data-ttu-id="e2713-114">Remplacez le code existant par le script suivant, puis appuyez sur **exécuter**.</span><span class="sxs-lookup"><span data-stu-id="e2713-114">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="e2713-115">Cette opération permet de configurer le classeur avec des noms de feuille de calcul, de tableau et de tableau croisé dynamique cohérents.</span><span class="sxs-lookup"><span data-stu-id="e2713-115">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("SubjectPivot");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script-for-your-automated-workflow"></a><span data-ttu-id="e2713-116">Créer un script Office pour votre flux de travail automatisé</span><span class="sxs-lookup"><span data-stu-id="e2713-116">Create an Office Script for your automated workflow</span></span>

<span data-ttu-id="e2713-117">Nous allons créer un script qui enregistre des informations à partir d’un message électronique.</span><span class="sxs-lookup"><span data-stu-id="e2713-117">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="e2713-118">Nous souhaitons savoir comment les jours de la semaine où nous recevons le plus de courrier et le nombre d’expéditeurs uniques qui envoient ce message.</span><span class="sxs-lookup"><span data-stu-id="e2713-118">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="e2713-119">Notre classeur contient un tableau contenant les colonnes **Date**, **jour de la semaine**, **adresse e-mail**et **objet** .</span><span class="sxs-lookup"><span data-stu-id="e2713-119">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="e2713-120">Notre feuille de calcul contient également un tableau croisé dynamique qui fait pivoter le **jour de la semaine** et l' **adresse de messagerie** (il s’agit des hiérarchies de lignes).</span><span class="sxs-lookup"><span data-stu-id="e2713-120">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="e2713-121">Le nombre de **sujets** uniques est l’affichage des informations agrégées (hiérarchie des données).</span><span class="sxs-lookup"><span data-stu-id="e2713-121">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="e2713-122">Notre script actualisera ce tableau croisé dynamique après la mise à jour de la table de messagerie.</span><span class="sxs-lookup"><span data-stu-id="e2713-122">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="e2713-123">Dans l' **éditeur de code**, sélectionnez **nouveau script**.</span><span class="sxs-lookup"><span data-stu-id="e2713-123">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="e2713-124">Le flux que nous allons créer plus tard dans le didacticiel enverra des informations de script sur chaque message électronique reçu.</span><span class="sxs-lookup"><span data-stu-id="e2713-124">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="e2713-125">Le script doit accepter cette entrée par le biais de paramètres dans la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="e2713-125">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="e2713-126">Remplacez le script par défaut par le script suivant :</span><span class="sxs-lookup"><span data-stu-id="e2713-126">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="e2713-127">Le script a besoin d’accéder à la table et au tableau croisé dynamique du classeur.</span><span class="sxs-lookup"><span data-stu-id="e2713-127">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="e2713-128">Ajoutez le code suivant au corps du script, après l’ouverture `{` :</span><span class="sxs-lookup"><span data-stu-id="e2713-128">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="e2713-129">Le `dateReceived` paramètre est de type `string` .</span><span class="sxs-lookup"><span data-stu-id="e2713-129">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="e2713-130">Nous allons convertir cela en [ `Date` objet](../develop/javascript-objects.md#date) afin que nous puissions facilement obtenir le jour de la semaine.</span><span class="sxs-lookup"><span data-stu-id="e2713-130">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="e2713-131">Une fois cette opération effectuée, nous devons mapper la valeur de nombre du jour à une version plus lisible.</span><span class="sxs-lookup"><span data-stu-id="e2713-131">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="e2713-132">Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :</span><span class="sxs-lookup"><span data-stu-id="e2713-132">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Parse the received date string.
    let date = new Date(dateReceived);

    // Convert number representing the day of the week into the name of the day.
    let dayText : string;
    switch (date.getDay()) {
      case 0:
        dayText = "Sunday";
        break;
      case 1:
        dayText = "Monday";
        break;
      case 2:
        dayText = "Tuesday";
        break;
      case 3:
        dayText = "Wednesday";
        break;
      case 4:
        dayText = "Thursday";
        break;
      case 5:
        dayText = "Friday";
        break;
      default:
        dayText = "Saturday";
        break;
    }
    ```

5. <span data-ttu-id="e2713-133">La `subject` chaîne peut inclure la balise de réponse « re : ».</span><span class="sxs-lookup"><span data-stu-id="e2713-133">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="e2713-134">Nous allons supprimer cela de la chaîne afin que les courriers électroniques dans le même thread aient le même objet pour le tableau.</span><span class="sxs-lookup"><span data-stu-id="e2713-134">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="e2713-135">Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :</span><span class="sxs-lookup"><span data-stu-id="e2713-135">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="e2713-136">Maintenant que les données de messagerie ont été mises en forme à notre gré, nous allons ajouter une ligne à la table de messagerie.</span><span class="sxs-lookup"><span data-stu-id="e2713-136">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="e2713-137">Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :</span><span class="sxs-lookup"><span data-stu-id="e2713-137">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. <span data-ttu-id="e2713-138">Enfin, nous allons nous assurer que le tableau croisé dynamique est actualisé.</span><span class="sxs-lookup"><span data-stu-id="e2713-138">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="e2713-139">Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :</span><span class="sxs-lookup"><span data-stu-id="e2713-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="e2713-140">Renommez votre **courrier électronique enregistrer** le script et appuyez sur **enregistrer le script**.</span><span class="sxs-lookup"><span data-stu-id="e2713-140">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="e2713-141">Votre script est maintenant prêt pour un flux de travail Automated Power.</span><span class="sxs-lookup"><span data-stu-id="e2713-141">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="e2713-142">Il doit ressembler au script suivant :</span><span class="sxs-lookup"><span data-stu-id="e2713-142">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Pivot");
  let pivotTable = pivotTableWorksheet.getPivotTable("SubjectPivot");

  // Parse the received date string.
  let date = new Date(dateReceived);

  // Convert number representing the day of the week into the name of the day.
  let dayText: string;
  switch (date.getDay()) {
    case 0:
      dayText = "Sunday";
      break;
    case 1:
      dayText = "Monday";
      break;
    case 2:
      dayText = "Tuesday";
      break;
    case 3:
      dayText = "Wednesday";
      break;
    case 4:
      dayText = "Thursday";
      break;
    case 5:
      dayText = "Friday";
      break;
    default:
      dayText = "Saturday";
      break;
  }

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayText, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="e2713-143">Créer un flux de travail automatisé avec Power automate</span><span class="sxs-lookup"><span data-stu-id="e2713-143">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="e2713-144">Connectez-vous au [site d’automate d’automate Power](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="e2713-144">Sign in to the [Power Automate preview site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="e2713-145">Dans le menu affiché sur le côté gauche de l’écran, appuyez sur **créer**.</span><span class="sxs-lookup"><span data-stu-id="e2713-145">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="e2713-146">Cela vous permet de créer de nouveaux flux de travail.</span><span class="sxs-lookup"><span data-stu-id="e2713-146">This brings you to list of ways to create new workflows.</span></span>

    ![Bouton créer dans Power automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="e2713-148">Dans la section **commencer à partir d’un champ vide** , sélectionnez **flux automatisé**.</span><span class="sxs-lookup"><span data-stu-id="e2713-148">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="e2713-149">Cette méthode crée un flux de travail déclenché par un événement, comme la réception d’un message électronique.</span><span class="sxs-lookup"><span data-stu-id="e2713-149">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![Option de flux automatisée dans Power Automated.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="e2713-151">Dans la fenêtre de boîte de dialogue qui s’affiche, entrez un nom pour votre flux dans la zone de texte **nom du flux** .</span><span class="sxs-lookup"><span data-stu-id="e2713-151">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="e2713-152">Ensuite, sélectionnez **quand un nouveau courrier électronique est reçu** dans la liste des options sous **choisir le déclencheur de votre flux**.</span><span class="sxs-lookup"><span data-stu-id="e2713-152">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="e2713-153">Vous devrez peut-être Rechercher l’option à l’aide de la zone de recherche.</span><span class="sxs-lookup"><span data-stu-id="e2713-153">You may need to search for the option using the search box.</span></span> <span data-ttu-id="e2713-154">Enfin, appuyez sur **créer**.</span><span class="sxs-lookup"><span data-stu-id="e2713-154">Finally, press **Create**.</span></span>

    ![Partie de la fenêtre créer un flux automatique dans Power automate, qui affiche l’option « nouveau courrier électronique ».](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="e2713-156">Ce didacticiel utilise Outlook.</span><span class="sxs-lookup"><span data-stu-id="e2713-156">This tutorial uses Outlook.</span></span> <span data-ttu-id="e2713-157">N’hésitez pas à utiliser votre service de messagerie préféré, bien que certaines options soient différentes.</span><span class="sxs-lookup"><span data-stu-id="e2713-157">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="e2713-158">Appuyez sur **nouvelle étape**.</span><span class="sxs-lookup"><span data-stu-id="e2713-158">Press **New step**.</span></span>

6. <span data-ttu-id="e2713-159">Sélectionnez l’onglet **standard** , puis **Excel Online (professionnel)**.</span><span class="sxs-lookup"><span data-stu-id="e2713-159">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Option Power automate pour Excel Online (professionnel).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="e2713-161">Sous **actions**, sélectionnez **exécuter un script (aperçu)**.</span><span class="sxs-lookup"><span data-stu-id="e2713-161">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Option d’action automate Power pour exécuter un script (aperçu).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="e2713-163">Spécifiez les paramètres suivants pour le connecteur de **script d’exécution** :</span><span class="sxs-lookup"><span data-stu-id="e2713-163">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="e2713-164">**Emplacement**: OneDrive entreprise</span><span class="sxs-lookup"><span data-stu-id="e2713-164">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="e2713-165">**Bibliothèque de documents**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="e2713-165">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="e2713-166">**Fichier**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="e2713-166">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="e2713-167">**Script**: enregistrer le courrier électronique</span><span class="sxs-lookup"><span data-stu-id="e2713-167">**Script**: Record Email</span></span>
    - <span data-ttu-id="e2713-168">**from**: from *(contenu dynamique d’Outlook)*</span><span class="sxs-lookup"><span data-stu-id="e2713-168">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="e2713-169">**dateReceived**: heure *de réception (contenu dynamique d’Outlook)*</span><span class="sxs-lookup"><span data-stu-id="e2713-169">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="e2713-170">**Subject**: subject *(contenu dynamique d’Outlook)*</span><span class="sxs-lookup"><span data-stu-id="e2713-170">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="e2713-171">*Notez que les paramètres du script s’affichent uniquement une fois que le script est sélectionné.*</span><span class="sxs-lookup"><span data-stu-id="e2713-171">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![Option d’action automate Power pour exécuter un script (aperçu).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="e2713-173">Cliquez sur **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="e2713-173">Press **Save**.</span></span>

<span data-ttu-id="e2713-174">Votre flux est maintenant activé.</span><span class="sxs-lookup"><span data-stu-id="e2713-174">Your flow is now enabled.</span></span> <span data-ttu-id="e2713-175">Il exécute automatiquement votre script chaque fois que vous recevez un courrier électronique via Outlook.</span><span class="sxs-lookup"><span data-stu-id="e2713-175">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="e2713-176">Gérer le script dans Power automate</span><span class="sxs-lookup"><span data-stu-id="e2713-176">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="e2713-177">Dans la page principale de l’alimentation automatique, sélectionnez **mes flux**.</span><span class="sxs-lookup"><span data-stu-id="e2713-177">From the main Power Automate page, select **My flows**.</span></span>

    ![Bouton mes flux dans Power automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="e2713-179">Sélectionnez votre flux.</span><span class="sxs-lookup"><span data-stu-id="e2713-179">Select your flow.</span></span> <span data-ttu-id="e2713-180">Ici, vous pouvez voir l’historique d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e2713-180">Here you can see the run history.</span></span> <span data-ttu-id="e2713-181">Vous pouvez actualiser la page ou appuyer sur le bouton actualiser **toutes les exécutions** pour mettre à jour l’historique.</span><span class="sxs-lookup"><span data-stu-id="e2713-181">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="e2713-182">Le flux se déclenche peu après la réception d’un message électronique.</span><span class="sxs-lookup"><span data-stu-id="e2713-182">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="e2713-183">Testez le flux en envoyant votre courrier.</span><span class="sxs-lookup"><span data-stu-id="e2713-183">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="e2713-184">Lorsque le flux est déclenché et exécute correctement votre script, vous devez voir la table du classeur et la mise à jour du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="e2713-184">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![La table de messagerie une fois que le flux a été exécuté plusieurs fois.](../images/power-automate-params-tutorial-4.png)

![Le tableau croisé dynamique après le flux a été exécuté plusieurs fois.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="e2713-187">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="e2713-187">Next steps</span></span>

<span data-ttu-id="e2713-188">Consultez la rubrique [exécuter des scripts Office avec Power automate](../develop/power-automate-integration.md) pour en savoir plus sur la connexion de scripts Office avec Power Automated.</span><span class="sxs-lookup"><span data-stu-id="e2713-188">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="e2713-189">Vous pouvez également consulter le [scénario d’exemple de rappel de tâche automatisée](../resources/scenarios/task-reminders.md) pour savoir comment combiner des scripts Office et alimenter automatiquement avec des cartes adaptatives de teams.</span><span class="sxs-lookup"><span data-stu-id="e2713-189">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
