---
title: Exécuter un script sur tous les fichiers Excel d’un dossier
description: Découvrez comment exécuter un script sur tous les fichiers Excel dans un dossier sur OneDrive Entreprise.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: fb9a4deb01b52ef031cb1ba3400bd6f10de9d9f5
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545789"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Exécuter un script sur tous les fichiers Excel d’un dossier

Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise. Il peut également être utilisé sur un dossier SharePoint dossier.
Il effectue des calculs sur les fichiers Excel, ajoute le formatage, et insère un [commentaire qui @mentions un](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) collègue.

Téléchargez le <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true"> fichierhighlight-alert-excel-files.zip</a>, extraire les fichiers dans un dossier intitulé **Ventes utilisées** dans cet échantillon, et l’essayer vous-même!

## <a name="sample-code-add-formatting-and-insert-comment"></a>Exemple de code : Ajouter le formatage et insérer le commentaire

C’est le script qui s’exécute sur chaque cahier de travail individuel.

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

## <a name="power-automate-flow-run-the-script-on-every-workbook-in-the-folder"></a>Power Automate flux : exécutez le script sur chaque cahier de travail dans le dossier

Ce flux exécute le script sur chaque cahier de travail dans le dossier « Ventes ».

1. Créez un nouveau **flux cloud instantané**.
1. Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer**.
1. Ajoutez une **nouvelle étape qui** utilise le **connecteur OneDrive Entreprise** liste et les fichiers Liste dans **l’action du** dossier.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Le connecteur OneDrive Entreprise terminé en Power Automate":::
1. Sélectionnez le dossier « Ventes » avec les cahiers de travail extraits.
1. Pour vous assurer que seuls les cahiers de travail sont sélectionnés, **choisissez Nouvelle étape,** puis **sélectionnez Condition** et définissez les valeurs suivantes :
    1. **Nom** (la valeur OneDrive nom du fichier)
    1. « se termine par »
    1. « xlsx ».

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Le bloc Power Automate condition qui applique les actions ultérieures à chaque fichier":::
1. Sous la **branche Si oui,** ajoutez le **connecteur Excel en ligne (Business)** avec l’action **script Run.** Utilisez les valeurs suivantes pour l’action :
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **Fichier**: **Id** (la valeur d’identification OneDrive fichier)
    1. **Script**: Votre nom de script

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Le connecteur Excel en ligne (Business) terminé en Power Automate":::
1. Enregistrez le flux et essayez-le.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vidéo de formation : Exécutez un script sur tous les Excel fichiers dans un dossier

[Regardez Sudhi Ramamurthy marcher à travers cet échantillon sur YouTube](https://youtu.be/xMg711o7k6w).
