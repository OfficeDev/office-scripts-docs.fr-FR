---
title: Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement
description: Un didacticiel qui présente comment envoyer des e-mails de rappel en exécutant des scripts Office pour Excel sur le web via Power Automate.
ms.date: 12/15/2020
localization_priority: Priority
ms.openlocfilehash: e7f1051076bf84cfbbec0fcdd72777766dbcf152
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545001"
---
# <a name="return-data-from-a-script-to-an-automatically-run-power-automate-flow"></a>Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement

Ce tutoriel vous apprend à renvoyer les informations d’un script Office pour Excel sur le web en tant qu’élément du flux de travail automatisé [Power Automate](https://flow.microsoft.com). Vous créerez un script qui parcoure un planning et fonctionne avec un flux pour envoyer des courriers de rappel. Ce flux s’exécutera selon un calendrier régulier, fournissant ces rappels à votre place.

> [!TIP]
> Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md).
>
> Si vous débutez avec Power Automate, nous vous recommandons de démarrer par les didacticiels [Appeler des scripts à partir d’un flux manuel Power Automate](excel-power-automate-manual.md) et [Transmettre des données à des scripts dans un flux automatique Power Automate (Aperçu)](excel-power-automate-trigger.md).
>
> [Les scripts Office utilisent TypeScript](../overview/code-editor-environment.md), et ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript. Si vous découvrez JavaScript, nous vous conseillons de commencer par consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Configuration requise

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Préparer le classeur

1. Téléchargez le classeur <a href="on-call-rotation.xlsx">on-call-rotation.xlsx</a> dans votre OneDrive.

1. Ouvrez **on-call-rotation.xlsx** dans Excel sur le web.

1. Ajoutez une ligne au tableau avec votre nom, adresse e-mail et les dates de début et de fin qui chevauchent la date actuelle.

    > [!IMPORTANT]
    > Le script que vous écrivez utilise la première entrée correspondante dans le tableau. Vérifiez donc que votre nom figure au-dessus des lignes de la semaine actuelle.

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="Feuille de calcul contenant les données du tableau de rotation des astreintes":::

## <a name="create-an-office-script"></a>Créer un script Office

1. Accédez à l’onglet **Automatiser**, puis sélectionnez **Tous les scripts**.

1. Sélectionnez **Nouveau script**.

1. Nommez le script **Appeler la personne d’astreinte**.

1. Vous devez désormais avoir un script vide. Nous utilisons le script pour obtenir l’adresse e-mail à partir de la feuille de calcul. Modifiez `main` pour renvoyer une chaîne, comme suit :

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. Ensuite, nous devons obtenir toutes les données du tableau. Cela nous permet de parcourir chaque ligne avec le script. Ajoutez le code suivant à l’intérieur de la fonction`main`.

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. Les dates du tableau sont stockées en utilisant le [Numéro de série de la date d’Excel](https://support.microsoft.com/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487). Nous convertissons ces dates en dates JavaScript pour les comparer. Nous ajoutons une fonction d’assistance à notre script. Ajoutez le code suivant à l’extérieur de la fonction`main` :

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. Nous devons maintenant déterminer la personne d’astreinte en ce moment. Sa ligne possède une date de début et une date de fin entourant la date actuelle. Nous écrivons un script pour partir du principe qu’une seule personne à la fois est d’astreinte. Les scripts peuvent renvoyer des tableaux pour traiter plusieurs valeurs, mais pour l’instant, nous renvoyons la première adresse e-mail qui correspond. Ajoutez la fonction suivante à la fin de la fonction `main`.

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. La méthode finale doit ressembler à ce qui suit :

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## <a name="create-an-automated-workflow-with-power-automate"></a>Créer un flux de travail automatisé avec Power Automate

1. Connectez-vous au site [Power Automate](https://flow.microsoft.com).

1. Dans le menu qui s’affiche sur le côté gauche de l’écran, appuyez sur **Créer**. Cela affiche une liste des moyens de créer de nouveaux flux de travail.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Le bouton de Création dans Power Automate":::

1. Sous la section **Démarrer à partir de zéro**, sélectionnez **Flux cloud planifié**.

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text="Le bouton Flux cloud planifié dans Power Automate":::

1. Nous devons maintenant définir le planning pour ce flux. Notre feuille de calcul a une nouvelle activité d’astreinte démarrant chaque lundi lors du premier semestre de 2021. Définissons le flux à exécuter en premier le lundi matin. Utilisez les options suivantes pour configurer le flux à exécuter chaque semaine le lundi.

    - **Nom de flux** : Avertir la personne d’astreinte
    - **Début** : 04/01/21 à 01h00
    - **Répéter tous les** : 1 semaine
    - **Durant ces journées** : M

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="Options d’affichage de la boîte de dialogue « Créer un flux cloud planifié ». Les options incluent le nom du flux, l’ de début, la fréquence de répétition et, un jour de la semaine pour exécuter le flux":::

1. Appuyez sur **Créer**.

1. Appuyez sur **Nouvelle étape**.

1. Sélectionnez l’onglet **Standard**, puis sélectionnez **Excel Online (Business)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Option Excel Online (Business) dans Power Automate":::

1. Sous **Actions**, sélectionnez **Exécuter le script**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Option d’action Exécuter un script dans Power Automate":::

1. Vous allez ensuite sélectionner le classeur et le script à utiliser dans l’étape de flux. Utilisez le classeur **rotation-des-astreintes.xlsx** que vous avez créé dans votre OneDrive. Spécifiez les paramètres suivants pour le connecteur **Exécuter le script** :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier** : rotation-des-astreintes.xlsx *(choisi via l’Explorateur de fichiers)*
    - **Script** : Obtenir la personne d’astreinte

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="Paramètres du connecteur Power Automate permettant d’exécuter un script":::

1. Appuyez sur **Nouvelle étape**.

1. Nous allons terminer le flux en envoyant un e-mail de rappel. Sélectionnez **Envoyer un e-mail (V2)** en utilisant la barre de recherche du connecteur. Utilisez le contrôle **Ajouter du contenu dynamique** pour ajouter l’adresse e-mail renvoyée par le script. Cette action va étiqueter **résultat** avec l’icône Excel à côté. Vous pouvez fournir tout objet et corps de texte de votre choix.

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="Les paramètres du connecteur Power Automate Outlook pour l’envoi d’un e-mail. Les options incluent le fichier à envoyer, l’objet de l’e-mail, le corps de l’e-mail, ainsi que des options avancées":::

    > [!NOTE]
    > Ce didacticiel utilise Outlook. N’hésitez pas à utiliser votre service de messagerie préféré, même si certaines options peuvent être différentes.

1. Appuyez sur **Enregistrer**.

## <a name="test-the-script-in-power-automate"></a>Tester le script dans Power Automate

Votre flux va s’exécuter chaque lundi matin. Vous pouvez tester le script maintenant en appuyant sur le bouton **Test** dans le coin supérieur droit de l’écran. Sélectionnez **Manuellement** et appuyez sur **Exécuter le test** pour exécuter le flux maintenant et tester le comportement. Vous devrez peut-être octroyer des autorisations à Excel et Outlook pour continuer.

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="Le bouton de Test de Power Automate":::

> [!TIP]
> Si votre flux ne parvient pas à envoyer un e-mail, revérifiez dans la feuille de calcul qu’une adresse e-mail valide figure dans la plage de dates actuelle en haut du tableau.

## <a name="next-steps"></a>Étapes suivantes

Visitez [Exécuter des scripts Office avec Power Automate](../develop/power-automate-integration.md) pour en savoir plus sur la connexion de scripts Office avec Power Automate.

Vous pouvez également consulter le [scénario type des rappels de tâches automatisés](../resources/scenarios/task-reminders.md) pour découvrir comment combiner les scripts Office et Power Automate avec les cartes adaptatives Teams.
