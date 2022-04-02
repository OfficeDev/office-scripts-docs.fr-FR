---
title: Appeler des scripts à partir d’un flux manuel Power Automate
description: Un tutoriel sur l’utilisation des scripts Office dans Power Automate via un déclencheur manuel.
ms.date: 06/29/2021
ms.localizationpriority: high
ms.openlocfilehash: e926540976dc066b3f07620c1e710dfa3abc7660
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585939"
---
# <a name="call-scripts-from-a-manual-power-automate-flow"></a>Appeler des scripts à partir d’un flux manuel Power Automate

Ce tutoriel vous apprend à exécuter un script Office pour Excel sur le web via [Power Automate](https://flow.microsoft.com). Vous allez créer un script qui met à jour les valeurs de deux cellules en y indiquant la date et l’heure de son exécution. Vous allez ensuite connecter ce script à un flux Power Automate déclenché manuellement, pour que le script s’exécute à chaque fois qu’un bouton est sélectionné dans Power Automate. Après avoir assimilé le modèle de base, vous pourrez développer le flux pour inclure d’autres applications et automatiser davantage votre flux de travail quotidien.

> [!TIP]
> Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md). [Les scripts Office utilisent TypeScript](../overview/code-editor-environment.md), et ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript. Si vous découvrez JavaScript, nous vous conseillons de commencer par consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## <a name="prerequisites"></a>Configuration requise

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## <a name="prepare-the-workbook"></a>Préparer le classeur

Power Automate ne peut pas utiliser de [références relatives](../testing/power-automate-troubleshooting.md#avoid-relative-references) comme `Workbook.getActiveWorksheet`pour accéder aux composants du classeur. Nous avons donc besoin d’un classeur et d’une feuille de calcul avec des noms cohérents que Power Automate peut référencer.

1. Créer un classeur nommé **MyWorkbook**.

2. Dans le classeur **MyWorkbook**, créez une feuille de calcul appelée **TutorialWorksheet**.

## <a name="create-an-office-script"></a>Créer un script Office

1. Accédez à l’onglet **Automatiser**, puis sélectionnez **Tous les scripts**.

2. Sélectionnez **Nouveau script**.

3. Remplacez le script par défaut par le script suivant. Ce script ajoute la date et l’heure actuelles aux deux premières cellules de la feuille de calcul **TutorialWorksheet**.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. Renommez le script **Définir la date et l’heure**. Sélectionnez le nom du script pour le modifier.

5. Enregistrez le script en sélectionnant **Enregistrer le script**.

## <a name="create-an-automated-workflow-with-power-automate"></a>Créer un flux de travail automatisé avec Power Automate

1. Connectez-vous au site [Power Automate](https://flow.microsoft.com).

2. Dans le menu qui s’affiche sur le côté gauche de l’écran, sélectionnez **Créer**. Cela affiche une liste des moyens de créer de nouveaux flux de travail.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="Bouton « Créer » de Power Automate":::

3. Dans la section **Démarrer à partir de zéro**, sélectionnez **Flux instantané**. Cela crée un flux de travail activé manuellement.

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="Option Flux instantané de Power Automate pour créer un nouveau flux de travail":::

4. Dans la fenêtre de dialogue qui s’affiche, entrez un nom pour votre flux dans la zone de texte **Nom du flux**, sélectionnez **Déclencher manuellement un flux** dans la liste des options sous **Choisir comment déclencher le flux**, puis sélectionnez **Créer**.

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="Option « Déclencher un flux manuellement » de Power Automate":::

    Notez qu’un flux déclenché manuellement n’est que l’un des nombreux types de flux. Dans le tutoriel suivant, vous allez créer un flux qui s’exécute automatiquement lorsque vous recevez un e-mail.

5. Sélectionnez **Nouvelle étape**.

6. Sélectionnez l’onglet **Standard**, puis sélectionnez **Excel Online (Business)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Option Excel en ligne (Business) dans Power Automate.":::

7. Sous **Actions**, sélectionnez **Exécuter le script**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Option Excel en ligne (Business) dans Power Automate. ":::

8. Vous allez ensuite sélectionner le classeur et le script à utiliser dans l’étape de flux. À titre de didacticiel, vous allez utiliser le classeur précédemment créé dans OneDrive, mais vous pouvez utiliser n’importe quel classeur dans un site OneDrive ou SharePoint. Spécifiez les paramètres suivants pour le connecteur **Exécuter le script** :

    - **Emplacement** : OneDrive Entreprise
    - **Bibliothèque de documents** : OneDrive
    - **Fichier** : MyWorkbook.xlsx *(choisi via l’Explorateur de fichiers)*
    - **Script** : Définir la date et l’heure

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="Paramètres du connecteur Power Automate permettant d’exécuter un script":::

9. Sélectionnez **Enregistrer**.

Votre flux est désormais prêt à être exécuté via Power Automate. Vous pouvez le tester à l’aide du bouton **Tester** dans l’éditeur de flux ou suivre les étapes restantes du tutoriel pour exécuter le flux à partir de votre collection de flux.

## <a name="run-the-script-through-power-automate"></a>Exécuter le script via Power Automate

1. Sur la page principale de Power Automate, sélectionnez **Mes flux**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="Bouton Mes flux de Power Automate":::

2. Sélectionnez **Mon flux de tutoriel** dans la liste des flux affichée dans l’onglet **Mes flux**. Cela affiche les informations sur le flux que nous avons créé précédemment.

3. Sélectionnez **Exécuter**.

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="Bouton Exécuter de Power Automate":::

4. Un volet des tâches apparaîtra pour exécuter le flux. Si vous êtes invité à vous **Connecter** à Excel Online, faites-le en sélectionnant **Continuer**.

5. Sélectionnez **Exécuter le flux**. Cela exécute le flux, qui exécute le script Office associé.

6. Sélectionnez **Terminé**. Vous devriez voir la section **Exécutions** s’actualiser en conséquence.

7. Actualisez la page pour voir les résultats de Power Automate. Si l’opération est réussie, accédez au classeur pour voir les cellules mises à jour. Si l’opération a échoué, vérifiez les paramètres du flux et exécutez-le une deuxième fois.

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="Sortie de Power Automate montrant une exécution réussie du flux":::

## <a name="next-steps"></a>Étapes suivantes

Suivez le tutoriel [Transférer des données aux scripts dans un flux Power Automate exécuté automatiquement](excel-power-automate-trigger.md). Il vous explique comment transmettre les données d’un service de flux de travail à votre script Office et comment exécuter le flux Power Automate lorsque certains événements se produisent.
