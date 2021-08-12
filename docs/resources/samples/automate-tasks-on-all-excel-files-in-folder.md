---
title: Exécuter un script sur tous les fichiers Excel d’un dossier
description: Découvrez comment exécuter un script sur tous les fichiers Excel dans un dossier sur OneDrive Entreprise.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: fad1483fbcddaf541874630e8a4e5a06faa784627d44d17ea2ab7ca0af1550a4
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847403"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Exécuter un script sur tous les fichiers Excel d’un dossier

Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise. Il peut également être utilisé sur un SharePoint dossier.
Il effectue des calculs sur les fichiers Excel, ajoute une mise en forme et insère un [commentaire](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) qui @mentions collègue.

## <a name="sample-excel-files"></a>Exemples Excel fichiers

Téléchargez <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a> tous les workbooks dont vous aurez besoin pour cet exemple. Extrayer ces fichiers dans un dossier intitulé **Ventes**. Ajoutez le script suivant à votre collection de scripts pour essayer l’exemple vous-même !

## <a name="sample-code-add-formatting-and-insert-comment"></a>Exemple de code : ajouter une mise en forme et insérer un commentaire

Il s’agit du script qui s’exécute sur chaque workbook individuel.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the table named "Table1" in the workbook.
  let table1 = workbook.getTable("Table1");

  // If the table is empty, end the script.
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }

  // Force the workbook to be completely recalculated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  // Get the "Amount Due" column from the table.
  const amountDueColumn = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueColumn.getRangeBetweenHeaderAndTotal().getValues();

  // Find the highest amount that's due.
  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }

  let highestAmountDue = table1.getColumn("Amount due").getRangeBetweenHeaderAndTotal().getRow(row);

  // Set the fill color to yellow for the cell with the highest value in the "Amount Due" column.
  highestAmountDue
    .getFormat()
    .getFill()
    .setColor("FFFF00");

  // Insert an @mention comment in the cell.
  workbook.addComment(highestAmountDue, {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate flux : exécuter le script sur chaque classeur du dossier

Ce flux exécute le script sur chaque classeur dans le dossier « Ventes ».

1. Créez un **flux de cloud instantané.**
1. Sélectionnez **Déclencher manuellement un flux,** puis **sélectionnez Créer.**
1. Ajoutez **une nouvelle étape qui** utilise le connecteur **OneDrive Entreprise** et les fichiers de liste **dans l’action de** dossier.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate.":::
1. Sélectionnez le dossier « Ventes » avec les classeurs extraits.
1. Pour vous assurer que seuls les workbooks sont sélectionnés, choisissez **Nouvelle étape,** puis **sélectionnez Condition**. Utilisez les valeurs suivantes pour la condition.
    1. **Nom** (valeur OneDrive nom de fichier)
    1. « se termine par »
    1. « xlsx »

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Le Power Automate condition qui applique les actions suivantes à chaque fichier.":::
1. Sous la **branche Si oui,** ajoutez **le connecteur Excel Online (Entreprise)** avec l’action **de script Exécuter.** Utilisez les valeurs suivantes pour l’action.
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **Fichier**: **ID** (valeur OneDrive’ID de fichier)
    1. **Script**: nom de votre script

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Le connecteur Excel Online (Entreprise) dans Power Automate.":::
1. Enregistrez le flux et testez-le. Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vidéo de formation : exécuter un script sur tous Excel fichiers d’un dossier

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/xMg711o7k6w).
