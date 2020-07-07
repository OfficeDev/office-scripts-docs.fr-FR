---
title: Exécuter automatiquement des scripts avec Power Automate
description: Didacticiel sur l’exécution de scripts Office pour Excel sur le Web via automate d’alimentation à l’aide d’un déclencheur externe automatique (réception de courriers électroniques via Outlook).
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: a750197d6b5ae770ad7d2e17b3ee00dc65ee8875
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043418"
---
# <a name="automatically-run-scripts-with-power-automate-preview"></a>Exécuter automatiquement des scripts avec Power automate (aperçu)

Ce didacticiel vous apprend à utiliser un script Office pour Excel sur le Web avec un flux de travail Automated [Power](https://flow.microsoft.com) . Votre script s’exécute automatiquement chaque fois que vous recevez un courrier électronique, en enregistrant des informations à partir du courrier électronique dans un classeur Excel.

## <a name="prerequisites"></a>Conditions préalables

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> Ce didacticiel suppose que vous avez terminé l' [exécution des scripts Office dans Excel sur le Web avec le didacticiel Power Automated](excel-power-automate-manual.md) .

## <a name="prepare-the-workbook"></a>Préparation du classeur

Power automate ne peut pas utiliser de [références relatives](../develop/power-automate-integration.md#avoid-using-relative-references) comme `Workbook.getActiveWorksheet` pour accéder aux composants du classeur. Par conséquent, nous avons besoin d’un classeur et d’une feuille de calcul avec des noms cohérents pour automate de puissance à référencer.

1. Créez un classeur nommé **MyWorkbook**.

2. Accédez à l’onglet **automatiser** et sélectionnez **éditeur de code**.

3. Sélectionnez **nouveau script**.

4. Remplacez le code existant par le script suivant, puis appuyez sur **exécuter**. Cette opération permet de configurer le classeur avec des noms de feuille de calcul, de tableau et de tableau croisé dynamique cohérents.

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

## <a name="create-an-office-script-for-your-automated-workflow"></a>Créer un script Office pour votre flux de travail automatisé

Nous allons créer un script qui enregistre des informations à partir d’un message électronique. Nous souhaitons savoir comment les jours de la semaine où nous recevons le plus de courrier et le nombre d’expéditeurs uniques qui envoient ce message. Notre classeur contient un tableau contenant les colonnes **Date**, **jour de la semaine**, **adresse e-mail**et **objet** . Notre feuille de calcul contient également un tableau croisé dynamique qui fait pivoter le **jour de la semaine** et l' **adresse de messagerie** (il s’agit des hiérarchies de lignes). Le nombre de **sujets** uniques est l’affichage des informations agrégées (hiérarchie des données). Notre script actualisera ce tableau croisé dynamique après la mise à jour de la table de messagerie.

1. Dans l' **éditeur de code**, sélectionnez **nouveau script**.

2. Le flux que nous allons créer plus tard dans le didacticiel enverra des informations de script sur chaque message électronique reçu. Le script doit accepter cette entrée par le biais de paramètres dans la `main` fonction. Remplacez le script par défaut par le script suivant :

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. Le script a besoin d’accéder à la table et au tableau croisé dynamique du classeur. Ajoutez le code suivant au corps du script, après l’ouverture `{` :

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. Le `dateReceived` paramètre est de type `string` . Nous allons convertir cela en [ `Date` objet](../develop/javascript-objects.md#date) afin que nous puissions facilement obtenir le jour de la semaine. Une fois cette opération effectuée, nous devons mapper la valeur de nombre du jour à une version plus lisible. Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :

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

5. La `subject` chaîne peut inclure la balise de réponse « re : ». Nous allons supprimer cela de la chaîne afin que les courriers électroniques dans le même thread aient le même objet pour le tableau. Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. Maintenant que les données de messagerie ont été mises en forme à notre gré, nous allons ajouter une ligne à la table de messagerie. Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. Enfin, nous allons nous assurer que le tableau croisé dynamique est actualisé. Ajoutez le code suivant à la fin de votre script, avant la fermeture `}` :

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Renommez votre **courrier électronique enregistrer** le script et appuyez sur **enregistrer le script**.

Votre script est maintenant prêt pour un flux de travail Automated Power. Il doit ressembler au script suivant :

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

## <a name="create-an-automated-workflow-with-power-automate"></a>Créer un flux de travail automatisé avec Power automate

1. Connectez-vous au [site d’automate d’automate Power](https://flow.microsoft.com).

2. Dans le menu affiché sur le côté gauche de l’écran, appuyez sur **créer**. Cela vous permet de créer de nouveaux flux de travail.

    ![Bouton créer dans Power automate.](../images/power-automate-tutorial-1.png)

3. Dans la section **commencer à partir d’un champ vide** , sélectionnez **flux automatisé**. Cette méthode crée un flux de travail déclenché par un événement, comme la réception d’un message électronique.

    ![Option de flux automatisée dans Power Automated.](../images/power-automate-params-tutorial-1.png)

4. Dans la fenêtre de boîte de dialogue qui s’affiche, entrez un nom pour votre flux dans la zone de texte **nom du flux** . Ensuite, sélectionnez **quand un nouveau courrier électronique est reçu** dans la liste des options sous **choisir le déclencheur de votre flux**. Vous devrez peut-être Rechercher l’option à l’aide de la zone de recherche. Enfin, appuyez sur **créer**.

    ![Partie de la fenêtre créer un flux automatique dans Power automate, qui affiche l’option « nouveau courrier électronique ».](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > Ce didacticiel utilise Outlook. N’hésitez pas à utiliser votre service de messagerie préféré, bien que certaines options soient différentes.

5. Appuyez sur **nouvelle étape**.

6. Sélectionnez l’onglet **standard** , puis **Excel Online (professionnel)**.

    ![Option Power automate pour Excel Online (professionnel).](../images/power-automate-tutorial-4.png)

7. Sous **actions**, sélectionnez **exécuter un script (aperçu)**.

    ![Option d’action automate Power pour exécuter un script (aperçu).](../images/power-automate-tutorial-5.png)

8. Spécifiez les paramètres suivants pour le connecteur de **script d’exécution** :

    - **Emplacement**: OneDrive entreprise
    - **Bibliothèque de documents**: OneDrive
    - **Fichier**: MyWorkbook.xlsx
    - **Script**: enregistrer le courrier électronique
    - **from**: from *(contenu dynamique d’Outlook)*
    - **dateReceived**: heure *de réception (contenu dynamique d’Outlook)*
    - **Subject**: subject *(contenu dynamique d’Outlook)*

    *Notez que les paramètres du script s’affichent uniquement une fois que le script est sélectionné.*

    ![Option d’action automate Power pour exécuter un script (aperçu).](../images/power-automate-params-tutorial-3.png)

9. Cliquez sur **Enregistrer**.

Votre flux est maintenant activé. Il exécute automatiquement votre script chaque fois que vous recevez un courrier électronique via Outlook.

## <a name="manage-the-script-in-power-automate"></a>Gérer le script dans Power automate

1. Dans la page principale de l’alimentation automatique, sélectionnez **mes flux**.

    ![Bouton mes flux dans Power automate.](../images/power-automate-tutorial-7.png)

2. Sélectionnez votre flux. Ici, vous pouvez voir l’historique d’exécution. Vous pouvez actualiser la page ou appuyer sur le bouton actualiser **toutes les exécutions** pour mettre à jour l’historique. Le flux se déclenche peu après la réception d’un message électronique. Testez le flux en envoyant votre courrier.

Lorsque le flux est déclenché et exécute correctement votre script, vous devez voir la table du classeur et la mise à jour du tableau croisé dynamique.

![La table de messagerie une fois que le flux a été exécuté plusieurs fois.](../images/power-automate-params-tutorial-4.png)

![Le tableau croisé dynamique après le flux a été exécuté plusieurs fois.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a>Étapes suivantes

Consultez la rubrique [exécuter des scripts Office avec Power automate](../develop/power-automate-integration.md) pour en savoir plus sur la connexion de scripts Office avec Power Automated.

Vous pouvez également consulter le [scénario d’exemple de rappel de tâche automatisée](../resources/scenarios/task-reminders.md) pour savoir comment combiner des scripts Office et alimenter automatiquement avec des cartes adaptatives de teams.
