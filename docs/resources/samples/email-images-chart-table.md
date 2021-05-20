---
title: Envoyez par courriel les images d’un Excel graphique et d’une table
description: Apprenez à utiliser les scripts Office les Power Automate pour extraire et envoyer par courriel les images d’un graphique et d Excel table.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 54b6b67a0f211f2dc6c881bab17ff23220619e6e
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545774"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Utilisez Office scripts et des Power Automate pour envoyer des images d’un graphique et d’une table

Cet exemple utilise Office scripts et Power Automate pour créer un graphique. Il envoie ensuite des images du graphique et de sa table de base.

## <a name="example-scenario"></a>Exemple de scénario

* Calculez pour obtenir les derniers résultats.
* Créer un graphique.
* Obtenez des images de tableau et de table.
* Envoyez les images par courriel Power Automate.

_Données d’entrée_

:::image type="content" source="../../images/input-data.png" alt-text="Une feuille de travail montrant un tableau des données d’entrée":::

_Graphique de sortie_

:::image type="content" source="../../images/chart-created.png" alt-text="Le graphique de colonne créé montrant le montant dû par le client":::

_Courriel qui a été reçu par Power Automate flux_

:::image type="content" source="../../images/email-received.png" alt-text="L’e-mail envoyé par le flux montrant Excel graphique intégré dans le corps":::

## <a name="solution"></a>Solution

Cette solution a deux parties :

1. [Un script Office pour calculer et extraire le graphique Excel la table](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Un flux Power Automate pour invoquer le script et envoyer les résultats par courriel. Par exemple, sur la façon de le faire, voir [Créer un flux de travail automatisé avec Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Code de l’échantillon : Calculer et extraire Excel graphique et la table

Le script suivant calcule et extrait un graphique Excel tableau et une table.

Téléchargez l’exemple <a href="email-chart-table.xlsx"> deemail-chart-table.xlsxet </a> utilisez-le avec ce script pour l’essayer vous-même!

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate : Envoyez un e-mail au graphique et aux images de table

Ce flux exécute le script et envoie les e-mails les images retournées.

1. Créez un nouveau **flux cloud instantané**.
1. Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer**.
1. Ajoutez une **nouvelle étape qui** utilise le **connecteur Excel en ligne (Business)** avec l’action **de script Run.** Utilisez les valeurs suivantes pour l’action :
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier**: Votre cahier de travail [(sélectionné avec le choix du fichier)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **Script**: Votre nom de script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Le connecteur Excel en ligne (Business) terminé en Power Automate":::
1. Cet exemple utilise Outlook client de messagerie. Vous pouvez utiliser n’importe quel connecteur de messagerie Power Automate supports, mais le reste des étapes supposent que vous avez choisi Outlook. Ajoutez une **nouvelle étape qui** utilise le **connecteur Office 365 Outlook et** l’action Envoyer et envoyer des **e-mails (V2).** Utilisez les valeurs suivantes pour l’action :
    * **À**: Votre compte de messagerie de test (ou e-mail personnel)
    * **Objet**: Veuillez examiner les données du rapport
    * Pour le **champ** Corps, sélectionnez « Code View » `</>` () et entrez ce qui suit:

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Le connecteur Office 365 Outlook terminé en Power Automate":::
1. Enregistrez le flux et essayez-le.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vidéo de formation : Extraire et envoyer par courriel des images du graphique et de la table

[Regardez Sudhi Ramamurthy marcher à travers cet échantillon sur YouTube](https://youtu.be/152GJyqc-Kw).
