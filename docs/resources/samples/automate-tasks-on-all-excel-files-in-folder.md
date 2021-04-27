---
title: Exécuter un script sur tous les fichiers Excel d’un dossier
description: Découvrez comment exécuter un script sur tous les fichiers Excel dans un dossier sur OneDrive Entreprise.
ms.date: 04/02/2021
localization_priority: Normal
ms.openlocfilehash: 6376dcac0eb36c04c2b60b2717d18cd730a0a8ee
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026855"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a>Exécuter un script sur tous les fichiers Excel d’un dossier

Ce projet effectue un ensemble de tâches d'automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise. Il peut également être utilisé sur un SharePoint dossier.
Il effectue des calculs sur les fichiers Excel, ajoute une mise en forme et insère un [commentaire](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) qui @mentions collègue.

Téléchargez le fichier <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip,</a>extrayez les fichiers dans un dossier intitulé **Ventes** utilisés dans cet exemple, et essayez-le vous-même !

## <a name="sample-code-add-formatting-and-insert-comment"></a>Exemple de code : ajouter une mise en forme et insérer un commentaire

Il s'agit du script qui s'exécute sur chaque workbook individuel.

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
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
1. Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**
1. Ajoutez **une nouvelle étape qui** utilise le connecteur **OneDrive Entreprise** et les fichiers de liste **dans l'action de** dossier.

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-1.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate.":::
1. Sélectionnez le dossier « Ventes » avec les classeurs extraits.
1. Pour vous assurer que seuls les workbooks sont sélectionnés, choisissez **Nouvelle étape,** puis **sélectionnez Condition** et définissez les valeurs suivantes :
    1. **Nom** (valeur OneDrive nom de fichier)
    1. « se termine par »
    1. « xlsx ».

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-2.png" alt-text="Le Power Automate condition qui applique les actions suivantes à chaque fichier.":::
1. Sous la **branche Si oui,** ajoutez **le connecteur Excel Online (Entreprise)** avec l'action **Exécuter le script (prévisualisation).** Utilisez les valeurs suivantes pour l'action :
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **File**: **ID** (valeur OneDrive'ID de fichier)
    1. **Script**: nom de votre script

    :::image type="content" source="../../images/all-files-in-folder-sample-flow-3.png" alt-text="Le connecteur Excel Online (Entreprise) dans Power Automate.":::
1. Enregistrez le flux et testez-le.

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a>Vidéo de formation : exécuter un script sur tous Excel fichiers d'un dossier

[Regardez une vidéo pas à pas](https://youtu.be/xMg711o7k6w) sur la façon d'exécuter un script sur tous les fichiers Excel dans un dossier OneDrive Entreprise ou SharePoint dossier.
