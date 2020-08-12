---
title: Transmettre des données à des scripts dans un flux automatique Power Automate
description: Un tutoriel sur l'exécution de scripts Office pour Excel sur le web via Power automate lorsque les messages sont reçus et transmettent les données de flux au script.
ms.date: 07/24/2020
localization_priority: Priority
ms.openlocfilehash: aed34f4b93bbe22768aab73d7a7264cc7d3c3ee6
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616765"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a>Transmettre des données à des scripts dans un flux automatique Power Automate (Aperçu)

Ce tutoriel vous apprend à exécuter un script Office pour Excel sur le web via un flux de travail automatisé [Power Automate](https://flow.microsoft.com). Votre script s’exécute automatiquement chaque fois que vous recevez un courrier électronique, enregistrant les informations du courrier électronique dans un classeur Excel. La possibilité de transférer des données d’autres applications dans un script Office offre une flexibilité et une liberté considérables dans vos processus automatisés.

> [!TIP]
> Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md). Si vous débutez avec Power Automate, nous vous recommandons de démarrer par le didacticiel [Appeler des scripts à partir d’un flux manuel Power Automate](excel-power-automate-manual.md). [Les scripts Office utilisent TypeScript](../overview/code-editor-environment.md), et ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript. Si vous débutez avec JavaScript, nous vous conseillons de démarrer par le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Configuration requise

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Préparer le classeur

Power Automate ne peut pas utiliser de [références relatives](../develop/power-automate-integration.md#avoid-using-relative-references) comme `Workbook.getActiveWorksheet`pour accéder aux composants du classeur. Nous avons donc besoin d’un classeur et d’une feuille de calcul avec des noms cohérents que Power Automate peut référencer.

1. Créer un nouveau classeur appelé **MyWorkbook**.

2. Accédez à l’onglet **Automatiser**, puis sélectionnez **Éditeur de code**.

3. Sélectionnez **Nouveau script**.

4. Remplacez le code existant par le script suivant et appuyez sur **Exécuter** : Cette opération permet de configurer le classeur avec des noms de feuille de calcul, de tableau et de tableau croisé dynamique cohérents.

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
      let pivotWorksheet = workbook.addWorksheet("Subjects");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script"></a>Créer un script Office

Créons un script qui enregistre les informations à partir d’un message électronique. Nous cherchons à identifier quels jours de la semaine nous recevons le plus de messages électroniques et combien d’expéditeurs uniques envoient ces messages électroniques. Notre classeur comporte une table avec les colonnes **date**, **jour de la semaine**, **adresse électronique**et **objet**. Notre feuille de calcul comporte également un tableau croisé dynamique qui fait pivoter le **jour de la semaine** et **adresse électronique** (il s’agit des hiérarchies de ligne). Le nombre de sujets **uniques** correspond aux informations agrégées affichées (hiérarchie des données). Notre script actualise ce tableau croisé dynamique après la mise à jour de la table de messagerie.

1. Dans l’ **Éditeur de code**, sélectionnez **Nouveau script**.

2. Le flux que nous allons créer plus tard dans le tutoriel enverra les informations de script de chaque message électronique reçu. Le script doit accepter cette entrée à l’aide de paramètres de la fonction `main`. Remplacez le script par défaut par le script suivant :

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. Le script a besoin d’accéder à la table et au tableau croisé dynamique du classeur. Ajoutez le code suivant dans le corps du script, après l'ouverture `{` :

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. Le paramètre `dateReceived` est de type `string`. Transformons cela en un [`Date` objet](../develop/javascript-objects.md#date) pour pouvoir obtenir facilement le jour de la semaine. Une fois cette opération effectuée, vous devez mapper la valeur numérique du jour à une version plus lisible. Ajoutez le code suivant à la fin de votre script (avant la clôture `}`) :

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. La chaîne de `subject` peut inclure la balise de réponse « RE : ». Supprimez-le de la chaîne afin que les messages électroniques d’un même fil de discussion aient le même objet pour le tableau. Ajoutez le code suivant à la fin de votre script (avant la clôture `}`) :

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. À présent que les données de courrier électronique ont été formatées à notre gout, ajoutons une ligne au tableau de courrier électronique. Ajoutez le code suivant à la fin de votre script (avant la clôture `}`) :

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. Enfin, assurez-vous que le tableau croisé dynamique est actualisé. Ajoutez le code suivant à la fin de votre script (avant la clôture `}`) :

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Renommez votre script **Enregistrer le courrier électronique**, puis appuyez sur **Enregistrer le script**.

Votre script est maintenant prêt pour un flux de travail Power Automate. Il devrait ressembler au script suivant :

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
  let pivotTableWorksheet = workbook.getWorksheet("Subjects");
  let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

  // Parse the received date string to determine the day of the week.
  let emailDate = new Date(dateReceived);
  let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a>Créer un flux de travail automatisé avec Power Automate

1. Connectez-vous au site [Power Automate](https://flow.microsoft.com).

2. Dans le menu qui s’affiche sur le côté gauche de l’écran, appuyez sur **Créer**. Cela affiche une liste des moyens de créer de nouveaux flux de travail.

    ![Le bouton Créer dans Power Automate.](../images/power-automate-tutorial-1.png)

3. Dans la section **Démarrer à partir de zéro**, sélectionnez **Flux automatique**. Cela permet de créer un flux de travail déclenché par un événement, par exemple, la réception d’un courrier électronique.

    ![Option flux automatisé dans Power Automate.](../images/power-automate-params-tutorial-1.png)

4. Dans la fenêtre de boîte de dialogue qui s’affiche, entrez un nom pour votre flux dans la zone de texte **Nom du flux**. Sélectionnez ensuite **À l'arrivée d'un nouveau courrier électronique** dans la liste d’options sous **Sélectionnez le déclencheur de votre flux**. Vous devrez peut-être rechercher l’option dans la zone de recherche. Enfin, appuyez sur **Créer**.

    ![Partie de la fenêtre Créer un flux automatisé dans Power automate affichant l’option « nouveau message électronique ».](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > Ce tutoriel utilise Outlook. N’hésitez pas à utiliser votre service de messagerie préféré, même si certaines options peuvent être différentes.

5. Appuyez sur **Nouvelle étape**.

6. Sélectionnez l’onglet **Standard**, puis sélectionnez **Excel Online (Business)**.

    ![L’option Power Automate pour Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. Sous **Actions**, sélectionnez **Exécuter le script** (Aperçu).

    ![L’option d’action Power Automate pour exécuter le script (Aperçu).](../images/power-automate-tutorial-5.png)

8. Vous devez ensuite sélectionner le classeur, le script, puis les arguments de saisie de script à utiliser dans l’étape de flux. À titre de didacticiel, vous allez utiliser le classeur précédemment créé dans OneDrive, mais vous pouvez utiliser n’importe quel classeur dans un site OneDrive ou SharePoint. Spécifiez les paramètres suivants pour le connecteur **Exécuter le script** :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier** : MyWorkbook.xlsx
    - **Script** : Enregistrer le courrier électronique
    - **à partir de**: de *(contenu dynamique d’Outlook)*
    - **date de réception**: heure de réception *(contenu dynamique d’Outlook)*
    - **objet**: Objet *(contenu dynamique d’Outlook)*

    *Notez que les paramètres du script s’affichent uniquement une fois le script sélectionné.*

    ![L’option d’action Power Automate pour exécuter le script (Aperçu).](../images/power-automate-params-tutorial-3.png)

9. Appuyez sur **Enregistrer**.

Votre flux est désormais activé. Il exécute automatiquement votre script chaque fois que vous recevez un courrier électronique via Outlook.

## <a name="manage-the-script-in-power-automate"></a>Gérer le script dans Power Automate

1. Sur la page principale de Power Automate, sélectionnez **Mes flux**.

    ![Le bouton Mes flux dans Power Automate.](../images/power-automate-tutorial-7.png)

2. Sélectionnez votre flux. Ici, vous pouvez voir l’historique d’exécution. Vous pouvez actualiser la page ou appuyer sur le bouton **Actualiser toutes les exécutions** pour mettre à jour l’historique. Le flux se déclenche peu après la réception d’un message électronique. Testez le flux en envoyant un courrier électronique.

Lorsque le flux est déclenché et exécute votre script correctement, la table du classeur et la mise à jour du tableau croisé dynamique doivent s’afficher.

![La table de messages après le flux s’exécute plusieurs fois.](../images/power-automate-params-tutorial-4.png)

![Le tableau croisé dynamique après le flux s’exécute plusieurs fois.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a>Étapes suivantes

Visitez [Exécuter des scripts Office avec Power Automate](../develop/power-automate-integration.md) pour en savoir plus sur la connexion de scripts Office avec Power Automate.

Vous pouvez également consulter le [scénario type des rappels de tâches automatisés](../resources/scenarios/task-reminders.md) pour découvrir comment combiner les scripts Office et Power Automate avec les cartes adaptatives Teams.
