---
title: 'Exemple de scénario de scripts Office : rappels de tâche automatisée'
description: Un exemple qui utilise Power automate et des cartes adaptatives automatise les rappels de tâches dans une feuille de calcul de gestion de projets.
ms.date: 06/09/2020
localization_priority: Normal
ms.openlocfilehash: f764c37dafdd964e9435d504770d10b1608428b8
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878807"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Exemple de scénario de scripts Office : rappels de tâche automatisée

Dans ce scénario, vous gérez un projet. Vous utilisez une feuille de calcul Excel pour suivre le statut de vos employés tous les mois. Vous avez souvent besoin de rappeler aux utilisateurs de remplir leur statut, de sorte que vous ayez décidé d’automatiser ce processus de rappel.

Vous allez créer un flux automatique d’alimentation pour les messages dont les champs d’État sont manquants et leur appliquer les réponses à la feuille de calcul. Pour ce faire, vous allez développer une paire de scripts pour gérer l’utilisation du classeur. Le premier script obtient une liste de personnes dont l’État est vide et le deuxième script ajoute une chaîne d’État à la ligne de droite. Vous utiliserez également des [cartes adaptatives](/microsoftteams/platform/task-modules-and-cards/what-are-cards) pour que les employés entrent leur état directement à partir de la notification.

## <a name="scripting-skills-covered"></a>Compétences en matière de script

- Créer des flux dans Power Automated
- Transmettre des données à des scripts
- Renvoyer des données à partir de scripts
- Cartes adaptatives de teams
- Tables

## <a name="prerequisites"></a>Conditions préalables

Ce scénario utilise [Power Automated](https://flow.microsoft.com) et [Microsoft teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Vous aurez besoin des deux éléments associés au compte que vous utilisez pour le développement de scripts Office. Pour un accès gratuit à un abonnement de développeur Microsoft afin de découvrir et d’utiliser ces applications, envisagez de participer au [programme de développement microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instructions de configuration

1. Téléchargez <a href="task-reminders.xlsx">task-reminders.xlsx</a> vers votre espace OneDrive.

2. Ouvrez le classeur dans Excel sur le Web.

3. Sous l’onglet **automatiser** , ouvrez l' **éditeur de code**.

4. Tout d’abord, nous avons besoin d’un script pour obtenir tous les employés ayant des rapports d’État manquants de la feuille de calcul. Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.

    ```typescript
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
          people.push({ name: row[NAME_INDEX], email: row[EMAIL_INDEX] });
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

5. Enregistrez le script sous le nom **Get People**.

6. Ensuite, nous avons besoin d’un deuxième script pour traiter les cartes de rapports d’État et placer les nouvelles informations dans la feuille de calcul. Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.

    ```typescript
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

7. Enregistrez le script sous le nom **Save Status**.

8. À présent, nous devons créer le flux. Ouvrez [Power automate](https://flow.microsoft.com/).

    > [!TIP]
    > Si vous n’avez pas créé de flux avant, consultez notre didacticiel [commencer à utiliser des scripts avec Power automate](../../tutorials/excel-power-automate-manual.md) pour apprendre les bases.

9. Créez un **flux de messagerie instantanée**.

10. Choisissez **déclencher manuellement un flux** à partir des options, puis appuyez sur **créer**.

11. Le flux doit appeler le script **Get People** pour obtenir tous les employés avec des champs d’État vides. Appuyez sur **nouvelle étape** et sélectionnez **Excel Online (professionnel)**. Sous **actions**, sélectionnez **exécuter un script (aperçu)**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement**: OneDrive entreprise
    - **Bibliothèque de documents**: OneDrive
    - **Fichier**: task-reminders.xlsx
    - **Script**: obtenir des personnes

    ![Étape du flux de script de la première exécution.](../../images/scenario-task-reminders-first-flow-step.png)

12. Ensuite, le flux doit traiter chaque employé dans le tableau renvoyé par le script. Appuyez sur **nouvelle étape** et sélectionnez **publier une carte adaptative pour un utilisateur de teams et attendez une réponse**.

13. Pour le champ **destinataire** , ajoutez le **courrier électronique** à partir du contenu dynamique (la sélection comportera le logo Excel). L’ajout de **courrier** entraîne l’enchaînement de l’étape du flux par un bloc **apply to each** . Cela signifie que le tableau sera parcouru par Power Automated.

14. L’envoi d’une carte adaptative nécessite que le JSON de la carte soit fourni comme **message**. Vous pouvez utiliser le [Concepteur de cartes adaptatives](https://adaptivecards.io/designer/) pour créer des cartes personnalisées. Pour cet exemple, utilisez le code JSON suivant.  

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

15. Renseignez les champs restants comme suit :

    - **Mettre à jour le message**: Merci d’avoir soumis votre rapport d’État. Votre réponse a été ajoutée à la feuille de calcul.
    - **Mise à jour**de la carte : Oui

16. Dans le bloc **apply to each** , après l' **envoi d’une carte adaptative à un utilisateur de teams et l’attente d’une réponse**, appuyez sur **Ajouter une action**. Sélectionnez **Excel Online (professionnel)**. Sous **actions**, sélectionnez **exécuter un script (aperçu)**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement**: OneDrive entreprise
    - **Bibliothèque de documents**: OneDrive
    - **Fichier**: task-reminders.xlsx
    - **Script**: enregistrer l’État
    - **senderEmail**: e-mail *(contenu dynamique d’Excel)*
    - **statusReportResponse**: Response *(contenu dynamique de Teams)*

    ![Étape d’application appliquer à chaque flux.](../../images/scenario-task-reminders-last-flow-step.png)

17. Enregistrez le flux.

## <a name="running-the-flow"></a>Exécution du flux

Pour tester le flux, assurez-vous que toutes les lignes de tableau dont le statut est vide utilisent une adresse de messagerie liée à un compte Teams (vous devez probablement utiliser votre propre adresse de messagerie lors des tests).

Vous pouvez sélectionner **test** dans le concepteur de flux ou exécuter le flux à partir de la page **mes flux** . Après avoir démarré le flux et accepté l’utilisation des connexions requises, vous devez recevoir une carte adaptative de Power automate via Teams. Une fois que vous avez rempli le champ d’État dans la carte, le flux se poursuit et met à jour la feuille de calcul avec l’état que vous avez fourni.

### <a name="before-running-the-flow"></a>Avant d’exécuter le flux

![Feuille de calcul avec un rapport d’état contenant une entrée d’État manquante.](../../images/scenario-task-reminders-spreadsheet-before.png)

### <a name="receiving-the-adaptive-card"></a>Réception de la carte adaptative

![Une carte adaptative dans teams demande à l’employé une mise à jour de l’État.](../../images/scenario-task-reminders-adaptive-card.png)

### <a name="after-running-the-flow"></a>Après l’exécution du flux

![Feuille de calcul avec un rapport d’État avec une entrée d’état de remplissage.](../../images/scenario-task-reminders-spreadsheet-after.png)
