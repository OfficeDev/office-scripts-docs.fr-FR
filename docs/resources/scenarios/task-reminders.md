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
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Scénario d’exemple de scripts : rappels de tâches automatisés

Dans ce scénario, vous gérez un projet. Vous utilisez une feuille Excel pour suivre l’état de vos employés chaque mois. Vous devez souvent rappeler aux gens de remplir leur statut, vous avez donc décidé d’automatiser ce processus de rappel.

Vous créerez un flux de Power Automate pour envoyer des messages aux personnes ayant des champs d’état manquants et appliquerez leurs réponses à la feuille de calcul. Pour ce faire, vous développerez une paire de scripts pour gérer le travail avec le cahier de travail. Le premier script reçoit une liste de personnes ayant des statuts vierges et le deuxième script ajoute une chaîne de statut à la bonne ligne. Vous utiliserez également les cartes [adaptatives Teams pour que](/microsoftteams/platform/task-modules-and-cards/what-are-cards) les employés saisiront leur statut directement à partir de la notification.

## <a name="scripting-skills-covered"></a>Compétences de script couvertes

- Créer des flux dans Power Automate
- Transmettre des données aux scripts
- Renvoyer les données des scripts
- Teams Cartes adaptatives
- Tables

## <a name="prerequisites"></a>Configuration requise

Ce scénario utilise [Power Automate](https://flow.microsoft.com) et [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Vous aurez besoin à la fois associé au compte que vous utilisez pour développer Office scripts. Pour accéder gratuitement à un abonnement Microsoft Developer pour en savoir plus sur ces applications et y travailler, envisagez de rejoindre [le Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instructions d’installation

1. Téléchargez <a href="task-reminders.xlsx">task-reminders.xlsx</a> sur votre OneDrive.

2. Ouvrez le cahier de travail en Excel sur le Web.

3. Sous **l’onglet Automate,** ouvrez **tous les scripts**.

4. Tout d’abord, nous avons besoin d’un script pour obtenir tous les employés avec des rapports d’état qui manquent à la feuille de calcul. Dans le volet de tâche de l’éditeur de **code,** **appuyez sur Nouveau Script** et coller le script suivant dans l’éditeur.

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

5. Enregistrer le script avec le nom **Get People**.

6. Ensuite, nous avons besoin d’un deuxième script pour traiter les bulletins d’état et mettre les nouvelles informations dans la feuille de calcul. Dans le volet de tâche de l’éditeur de **code,** **appuyez sur Nouveau Script** et coller le script suivant dans l’éditeur.

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

7. Enregistrez le script avec le nom **Enregistrer le statut**.

8. Maintenant, nous devons créer le flux. Ouvrez [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Si vous n’avez pas créé un flux avant, s’il vous plaît consulter notre [tutoriel Commencez à utiliser des scripts avec Power Automate](../../tutorials/excel-power-automate-manual.md) pour apprendre les bases.

9. Créez un nouveau **flux instantané**.

10. Choisissez **de déclencher manuellement un flux à partir** des options et appuyez sur **Créer**.

11. Le flux doit appeler le script **Get People pour** obtenir tous les employés avec des champs de statut vides. Appuyez **sur Nouvelle** étape et **sélectionnez Excel en ligne (Affaires)**. Sous **Actions**, sélectionnez **Script d’exécuter**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier**: task-reminders.xlsx *(Choisi par le navigateur de fichiers)*
    - **Script**: Get People

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Le flux Power Automate affichage de la première étape de flux de script s’exécutent":::

12. Ensuite, le flux doit traiter chaque employé dans le tableau retourné par le script. Appuyez **sur Nouvelle** étape **et sélectionnez Poster une carte adaptative à Teams utilisateur et attendre une réponse**.

13. Pour le **champ** Destinataire, ajoutez **l’e-mail** du contenu dynamique (la sélection aura le logo Excel par elle). **L’ajout** d’e-mail provoque l’étape de flux d’être **entouré par une application à** chaque bloc. Cela signifie que le tableau sera itéré par Power Automate.

14. L’envoi d’une carte adaptative exige que le JSON de la carte soit fourni sous forme de **message.** Vous pouvez utiliser le concepteur [de cartes adaptatives](https://adaptivecards.io/designer/) pour créer des cartes personnalisées. Pour cet échantillon, utilisez le JSON suivant.  

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

15. Remplissez les champs restants comme suit :

    - **Message de mise** à jour : Merci d’avoir soumis votre rapport d’état. Votre réponse a été ajoutée avec succès à la feuille de calcul.
    - **Devrait mettre à jour la** carte : Oui

16. Dans **l’apply à chaque** bloc, en suivant **la publication d’une carte adaptative à un utilisateur Teams et attendre une réponse, appuyez** **sur Ajouter une action**. Sélectionnez **Excel en ligne (Affaires)**. Sous **Actions**, sélectionnez **Script d’exécuter**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier**: task-reminders.xlsx *(Choisi par le navigateur de fichiers)*
    - **Script**: Enregistrer le statut
    - **senderEmail**: e-mail *(contenu dynamique de Excel)*
    - **statusReportResponse**: réponse *(contenu dynamique de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Le Power Automate de la circulation montrant l’application à chaque étape":::

17. Enregistrez le flux.

## <a name="running-the-flow"></a>Exécution du flux

Pour tester le flux, assurez-vous que toutes les lignes de table avec l’état vierge utilisent une adresse e-mail liée à un compte Teams (vous devriez probablement utiliser votre propre adresse e-mail lors des tests).

Vous pouvez sélectionner Test **à partir** du concepteur de flux, ou exécuter le flux à partir de la page **Mes flux.** Après avoir commencé le flux et accepté l’utilisation des connexions requises, vous devez recevoir une carte adaptative de Power Automate à Teams. Une fois que vous remplissez le champ d’état de la carte, le flux se poursuivra et mettra à jour la feuille de calcul avec l’état que vous fournissez.

### <a name="before-running-the-flow"></a>Avant d’exécuter le flux

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Une feuille de travail avec un rapport d’état contenant une entrée d’état manquante":::

### <a name="receiving-the-adaptive-card"></a>Réception de la carte adaptative

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Une carte adaptative en Teams à l’employé pour une mise à jour de statut":::

### <a name="after-running-the-flow"></a>Après avoir fait fonctionner le flux

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Une feuille de travail avec un rapport d’état avec une entrée de statut maintenant remplie":::
