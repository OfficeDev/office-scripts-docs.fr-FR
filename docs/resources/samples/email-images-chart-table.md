---
title: Email les images d’un graphique et d’un tableau Excel
description: Découvrez comment utiliser les scripts Office et Power Automate pour extraire et envoyer par e-mail les images d’un graphique et d’un tableau Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: dbf9135723a735321c99991d94f4b4387d800702
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572464"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Utiliser les scripts Office et Power Automate pour envoyer par e-mail des images d’un graphique et d’un tableau

Cet exemple utilise les scripts Office et Power Automate pour créer un graphique. Il envoie ensuite par e-mail des images du graphique et de sa table de base.

## <a name="example-scenario"></a>Exemple de scénario

* Calculer pour obtenir les derniers résultats.
* Créez un graphique.
* Obtenir des images de graphique et de tableau.
* Email les images avec Power Automate.

_Données d’entrée_

:::image type="content" source="../../images/input-data.png" alt-text="Feuille de calcul montrant une table de données d’entrée.":::

_Graphique de sortie_

:::image type="content" source="../../images/chart-created.png" alt-text="Histogramme créé montrant le montant dû par le client.":::

_Email reçues via le flux Power Automate_

:::image type="content" source="../../images/email-received.png" alt-text="E-mail envoyé par le flux montrant le graphique Excel incorporé dans le corps.":::

## <a name="solution"></a>Solution

Cette solution comprend deux parties :

1. [Un script Office pour calculer et extraire un graphique et un tableau Excel](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Flux Power Automate pour appeler le script et envoyer par e-mail les résultats. Pour obtenir un exemple sur la façon de procéder, consultez [Créer un flux de travail automatisé avec Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez [email-chart-table.xlsx](email-chart-table.xlsx) pour un classeur prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Exemple de code : Calculer et extraire le graphique et le tableau Excel

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {
  // Recalculate the workbook to ensure all tables and charts are updated.
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  // Get the data from the "InvoiceAmounts" table.
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  // Delete the "ChartSheet" worksheet if it's present, then recreate it.
  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');

  // Add the selected data to the new worksheet.
  const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length-1, selectColumns[0].length-1);
  targetRange.setValues(selectColumns);

  // Insert the chart on sheet 'ChartSheet' at cell "D1".
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');

  // Get images of the chart and table, then return them for a Power Automate flow.
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {chartImage, tableImage};
}

// The interface for table and chart images.
interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Flux Power Automate : Email les images de graphique et de tableau

Ce flux exécute le script et envoie par e-mail les images retournées.

1. Créez un **flux de cloud instantané**.
1. Choisissez **déclencher manuellement un flux** , puis **sélectionnez Créer**.
1. Ajoutez une **nouvelle étape** qui utilise le connecteur **Excel Online (Entreprise)** avec l’action **Exécuter le script** . Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier** : votre classeur ([sélectionné avec le sélecteur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script** : nom de votre script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Connecteur Excel Online (Entreprise) terminé dans Power Automate.":::
1. Cet exemple utilise Outlook comme client de messagerie. Vous pouvez utiliser n’importe quel connecteur de messagerie pris en charge par Power Automate, mais les autres étapes supposent que vous avez choisi Outlook. Ajoutez une **nouvelle étape** qui utilise le **connecteur Office 365 Outlook** et l’action **Envoyer et envoyer un e-mail (V2**). Utilisez les valeurs suivantes pour l’action.
    * **À** : votre compte de messagerie de test (ou e-mail personnel)
    * **Objet** : Veuillez consulter les données du rapport
    * Pour le champ **Corps** , sélectionnez « Affichage du code » (`</>`) et entrez les éléments suivants :

    ```HTML
    <p>Please review the following report data:<br>
    <br>
    Chart:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/chartImage']}"/>
    <br>
    Data:<br>
    <br>
    <img src="data:image/png;base64,@{outputs('Run_script')?['body/result/tableImage']}"/>
    <br>
    </p>
    ```

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="L’Office 365 connecteur Outlook terminé dans Power Automate.":::
1. Enregistrez le flux et essayez-le. Utilisez le bouton **Tester** dans la page de l’éditeur de flux ou exécutez le flux dans l’onglet **Mes flux** . Veillez à autoriser l’accès lorsque vous y êtes invité.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vidéo de formation : Extraire et envoyer par e-mail des images de graphique et de tableau

[Regardez Sudhi Ramamurthy parcourir cet exemple sur YouTube](https://youtu.be/152GJyqc-Kw).
