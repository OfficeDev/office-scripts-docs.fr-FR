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
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a>Office Exemple de scénario de scripts : rappels de tâches automatisés

Dans ce scénario, vous gérez un projet. Vous utilisez une feuille Excel pour suivre l’état de vos employés tous les mois. Vous devez souvent rappeler aux personnes de remplir leur statut. Vous avez donc décidé d’automatiser ce processus de rappel.

Vous allez créer un flux Power Automate message aux personnes dont les champs d’état sont manquants et appliquer leurs réponses à la feuille de calcul. Pour ce faire, vous allez développer une paire de scripts pour gérer l’utilisation du classer. Le premier script obtient la liste des personnes dont l’état est vide et le second ajoute une chaîne d’état à la ligne de droite. Vous utiliserez également des cartes [adaptatives Teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) pour que les employés entrent leur état directement à partir de la notification.

## <a name="scripting-skills-covered"></a>Compétences d’écriture de scripts couvertes

- Créer des flux dans Power Automate
- Transmettre des données à des scripts
- Renvoyer des données à partir de scripts
- Teams Cartes adaptatives
- Tables

## <a name="prerequisites"></a>Conditions préalables

Ce scénario utilise [Power Automate](https://flow.microsoft.com) et [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). Vous aurez besoin des deux associés au compte que vous utilisez pour développer Office Scripts. Pour un accès gratuit à un abonnement Microsoft Développeur pour en savoir plus sur ces applications et travailler avec celles-ci, envisagez de rejoindre le programme [Microsoft 365 développeur microsoft.](https://developer.microsoft.com/microsoft-365/dev-program)

## <a name="setup-instructions"></a>Instructions d’installation

1. Téléchargez <a href="task-reminders.xlsx">task-reminders.xlsx</a> sur votre OneDrive.

1. Ouvrez le Excel sur le Web.

1. Tout d’abord, nous avons besoin d’un script pour obtenir tous les employés dont les rapports d’état sont manquants dans la feuille de calcul. Sous **l’onglet Automatiser,** sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.

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

1. Enregistrez le script avec le nom **Get People**.

1. Ensuite, nous avons besoin d’un second script pour traiter les cartes de rapport d’état et placer les nouvelles informations dans la feuille de calcul. Dans le volet Des tâches de l’Éditeur de code, sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.

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

1. Enregistrez le script sous le nom **Enregistrer l’état**.

1. Maintenant, nous devons créer le flux. Ouvrez [Power Automate](https://flow.microsoft.com/).

    > [!TIP]
    > Si vous n’avez pas encore créé de flux, consultez notre didacticiel Commencez à utiliser des [scripts](../../tutorials/excel-power-automate-manual.md) Power Automate pour en savoir plus sur les bases.

1. Créez un **flux instantané.**

1. Sélectionnez **Déclencher manuellement un flux à** partir des options et sélectionnez **Créer.**

1. Le flux doit appeler le script **Obtenir des** personnes pour obtenir tous les employés avec des champs d’état vides. Sélectionnez **Nouvelle étape,** puis **sélectionnez Excel Online (Entreprise).** Sous **Actions**, sélectionnez **Exécuter le script**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier**: task-reminders.xlsx *(choisi via le navigateur de fichiers)*
    - **Script**: obtenir des personnes

    :::image type="content" source="../../images/scenario-task-reminders-first-flow-step.png" alt-text="Flux Power Automate montrant la première étape du flux de script d’exécuter.":::

1. Ensuite, le flux doit traiter chaque employé dans le tableau renvoyé par le script. Sélectionnez **Nouvelle étape,** puis choisissez Publier une carte adaptative à **un utilisateur Teams et attendre une réponse.**

1. Pour le **champ Destinataire,** ajoutez **le courrier** électronique à partir du contenu dynamique (la sélection Excel logo). **L’ajout d’un** message électronique entraîne le fait que l’étape du flux soit entourée d’une **application à chaque** bloc. Cela signifie que le tableau sera itéré par Power Automate.

1. L’envoi d’une carte adaptative nécessite que le JSON de la carte soit fourni en tant que **message.** Vous pouvez utiliser le Concepteur de [cartes adaptatives pour](https://adaptivecards.io/designer/) créer des cartes personnalisées. Pour cet exemple, utilisez le JSON suivant.  

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

1. Remplissez les champs restants comme suit :

    - **Message de mise à** jour : merci d’avoir envoyé votre rapport d’état. Votre réponse a été ajoutée avec succès à la feuille de calcul.
    - **Doit mettre à jour la carte**: Oui

1. In the **Apply to each** block, following the Post an Adaptive Card to a Teams user and wait for a **response**, select Add **an action**. Sélectionnez **Excel Online (Entreprise).** Sous **Actions**, sélectionnez **Exécuter le script**. Fournissez les entrées suivantes pour l’étape de flux :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier**: task-reminders.xlsx *(choisi via le navigateur de fichiers)*
    - **Script**: enregistrer l’état
    - **senderEmail**: courrier *électronique (contenu dynamique de Excel)*
    - **statusReportResponse**: réponse *(contenu dynamique de Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-last-flow-step.png" alt-text="Le Power Automate flux montrant l’application à chaque étape.":::

1. Enregistrez le flux.

## <a name="running-the-flow"></a>Exécution du flux

Pour tester le flux, assurez-vous que les lignes de tableau dont l’état est vide utilisent une adresse de messagerie liée à un compte Teams (vous devez probablement utiliser votre propre adresse e-mail lors du test). Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.

Vous devez recevoir une carte adaptative de Power Automate à Teams. Une fois que vous avez rempli le champ d’état dans la carte, le flux continue et met à jour la feuille de calcul avec l’état que vous fournissez.

### <a name="before-running-the-flow"></a>Avant d’exécution du flux

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="Feuille de calcul avec un rapport d’état contenant une entrée d’état manquante.":::

### <a name="receiving-the-adaptive-card"></a>Réception de la carte adaptative

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="Une carte adaptative Teams demande à l’employé une mise à jour de l’état.":::

### <a name="after-running-the-flow"></a>Après l’exécution du flux

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="Feuille de calcul avec un rapport d’état avec une entrée d’état maintenant remplie.":::
