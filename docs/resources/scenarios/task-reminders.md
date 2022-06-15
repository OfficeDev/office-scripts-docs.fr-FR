---
title: 'exemple de scénario de scripts Office : rappels de tâches automatisés'
description: Exemple qui utilise Power Automate et les cartes adaptatives automatisent les rappels de tâches dans une feuille de calcul de gestion de projet.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 08f3713210e83162f86d38bc8eb33d76bf8a7288
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088112"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>exemple de scénario de scripts Office : rappels de tâches automatisés

Dans ce scénario, vous gérez un projet. Vous utilisez une feuille de calcul Excel pour suivre l’état de vos employés tous les mois. Vous devez souvent rappeler aux utilisateurs de remplir leur état. Vous avez donc décidé d’automatiser ce processus de rappel.

Vous allez créer un flux Power Automate pour envoyer des messages aux personnes ayant des champs d’état manquants et appliquer leurs réponses à la feuille de calcul. Pour ce faire, vous allez développer une paire de scripts pour gérer l’utilisation du classeur. Le premier script obtient une liste de personnes avec des états vides et le deuxième ajoute une chaîne d’état à la ligne de droite. Vous utiliserez également [Teams cartes adaptatives](/microsoftteams/platform/task-modules-and-cards/what-are-cards) pour que les employés entrent leur statut directement à partir de la notification.

## <a name="scripting-skills-covered"></a>Compétences de script couvertes

- Créer des flux dans Power Automate
- Transmettre des données à des scripts
- Retourner des données à partir de scripts
- Teams cartes adaptatives
- Tables

## <a name="prerequisites"></a>Conditions préalables

Ce scénario utilise [Power Automate](https://flow.microsoft.com) et [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Vous devez être associé au compte que vous utilisez pour développer des scripts Office. Pour accéder gratuitement à un abonnement Microsoft Developer afin d’en savoir plus sur ces applications et de les utiliser, envisagez de rejoindre le [programme Microsoft 365 développeurs](https://developer.microsoft.com/microsoft-365/dev-program).

## <a name="setup-instructions"></a>Instructions d’installation

1. Téléchargez <a href="task-reminders.xlsx">task-reminders.xlsx</a> sur votre OneDrive.

1. Ouvrez le classeur dans Excel sur le Web.

1. Tout d’abord, nous avons besoin d’un script pour obtenir tous les employés dont les rapports d’état sont manquants dans la feuille de calcul. Sous l’onglet **Automatiser** , sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.

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

1. Enregistrez le script portant le nom **Get People**.

1. Ensuite, nous avons besoin d’un deuxième script pour traiter les cartes de rapport d’état et placer les nouvelles informations dans la feuille de calcul. Dans le volet Office Éditeur de code, sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.

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

1. Enregistrez le script portant le nom **Enregistrer l’état**.

1. Maintenant, nous devons créer le flux. Ouvrez [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Si vous n’avez pas encore créé de flux, consultez notre didacticiel [Commencer à utiliser des scripts avec Power Automate](../../tutorials/excel-power-automate-manual.md) pour en savoir plus sur les principes de base.

1. Créez un **flux instantané**.

1. Choisissez **déclencher manuellement un flux** dans les options, puis sélectionnez **Créer**.

1. Le flux doit appeler le script **Get People** pour obtenir tous les employés avec des champs d’état vides. Sélectionnez **Nouvelle étape**, puis **sélectionnez Excel Online (Entreprise).** Sous **Actions**, sélectionnez **Exécuter le script**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier** : task-reminders.xlsx *(choisi par le biais du navigateur de fichiers)*
    - **Script** : Obtenir des personnes

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Flux Power Automate montrant la première étape du flux de script d’exécution.":::

1. Ensuite, le flux doit traiter chaque employé dans le tableau retourné par le script. Sélectionnez **Nouvelle étape**, puis **postez une carte adaptative à un utilisateur Teams et attendez une réponse**.

1. Pour le champ **Destinataire**, ajoutez un **e-mail** à partir du contenu dynamique (la sélection comportera le logo Excel). L’ajout **d’un e-mail** entraîne l’encerclement de l’étape de flux par une **application à chaque** bloc. Cela signifie que le tableau sera itéré par Power Automate.

1. L’envoi d’une carte adaptative nécessite que le [JSON](https://www.w3schools.com/whatis/whatis_json.asp) de la carte soit fourni en tant que **message**. Vous pouvez utiliser le [Concepteur de cartes adaptatives](https://adaptivecards.io/designer/) pour créer des cartes personnalisées. Pour cet exemple, utilisez le code JSON suivant.  

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

1. Renseignez les champs restants comme suit :

    - **Message de mise à jour** : Merci d’avoir envoyé votre rapport d’état. Votre réponse a été ajoutée à la feuille de calcul.
    - **Doit mettre à jour la carte** : Oui

1. Dans appliquer **à chaque** bloc, après **avoir posté une carte adaptative à un utilisateur Teams et attendre une réponse**, **sélectionnez Ajouter une action**. Sélectionnez **Excel Online (Entreprise).** Sous **Actions**, sélectionnez **Exécuter le script**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier** : task-reminders.xlsx *(choisi par le biais du navigateur de fichiers)*
    - **Script** : Enregistrer l’état
    - **senderEmail** : e-mail *(contenu dynamique de Excel)*
    - **statusReportResponse** : réponse *(contenu dynamique de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Flux Power Automate montrant l’étape Appliquer à chaque étape.":::

1. Enregistrez le flux.

## <a name="running-the-flow"></a>Exécution du flux

Pour tester le flux, assurez-vous que toutes les lignes de table avec un état vide utilisent une adresse e-mail liée à un compte Teams (vous devez probablement utiliser votre propre adresse e-mail lors du test). Utilisez le bouton **Tester** dans la page de l’éditeur de flux ou exécutez le flux dans l’onglet **Mes flux** . Veillez à autoriser l’accès lorsque vous y êtes invité.

Vous devez recevoir une carte adaptative de Power Automate à Teams. Une fois que vous avez renseigné le champ d’état dans la carte, le flux continue et met à jour la feuille de calcul avec l’état que vous fournissez.

### <a name="before-running-the-flow"></a>Avant d’exécuter le flux

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Feuille de calcul avec un rapport d’état contenant une entrée d’état manquante.":::

### <a name="receiving-the-adaptive-card"></a>Réception de la carte adaptative

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Une carte adaptative dans Teams demandant à l’employé une mise à jour de l’état.":::

### <a name="after-running-the-flow"></a>Après avoir exécuté le flux

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Feuille de calcul avec un rapport d’état avec une entrée d’état maintenant remplie.":::
