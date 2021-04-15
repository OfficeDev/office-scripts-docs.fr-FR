---
title: Envoyer par courrier électronique les images d'un graphique et d'un tableau Excel
description: Découvrez comment utiliser Office Scripts et Power Automate pour extraire et envoyer par courrier électronique les images d'un graphique et d'un tableau Excel.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: de3cf16537cb12db45d4d465d367d797d053afc4
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754809"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Utiliser Office Scripts et Power Automate pour envoyer des images électroniques d'un graphique et d'un tableau

Cet exemple utilise Office Scripts et Power Automate pour créer un graphique. Il envoie ensuite des images du graphique et de sa table de base par courrier électronique.

## <a name="example-scenario"></a>Exemple de scénario

* Calculer pour obtenir les derniers résultats.
* Créez un graphique.
* Obtenir des images de graphique et de tableau.
* Envoyez un e-mail aux images avec Power Automate.

_Données d'entrée_

:::image type="content" source="../../images/input-data.png" alt-text="Feuille de calcul montrant une table des données d'entrée.":::

_Graphique de sortie_

:::image type="content" source="../../images/chart-created.png" alt-text="Graphique en colonnes créé montrant le montant dû par le client.":::

_Courrier électronique reçu via le flux Power Automate_

:::image type="content" source="../../images/email-received.png" alt-text="Courrier électronique envoyé par le flux montrant le graphique Excel incorporé dans le corps.":::

## <a name="solution"></a>Solution

Cette solution est en deux parties :

1. [Script Office pour calculer et extraire un graphique et un tableau Excel](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Flux Power Automate pour appeler le script et envoyer par courrier électronique les résultats. Pour obtenir un exemple sur la procédure à suivre, voir Créer un flux de travail [automatisé avec Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Exemple de code : calculer et extraire un graphique et un tableau Excel

Le script suivant calcule et extrait un graphique et un tableau Excel.

Téléchargez l'exemple <a href="email-chart-table.xlsx">email-chart-table.xlsx</a> fichier et utilisez-le avec ce script pour l'essayer vous-même !

```TypeScript
function main(workbook: ExcelScript.Workbook): ReportImages {

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  
  let sheet1 = workbook.getWorksheet("Sheet1");
  const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
  const rows = table.getRange().getTexts();

  const selectColumns = rows.map((row) => {
    return [row[2], row[5]];
  });
  table.setShowTotals(true);
  selectColumns.splice(selectColumns.length-1, 1);
  console.log(selectColumns);

  workbook.getWorksheet('ChartSheet')?.delete();
  const chartSheet = workbook.addWorksheet('ChartSheet');
  const targetRange = updateRange(chartSheet, selectColumns);

  // Insert chart on sheet 'Sheet1'.
  let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
  chart_2.setPosition('D1');
  const chartImage = chart_2.getImage();
  const tableImage = table.getRange().getImage();
  return {
    chartImage,
    tableImage
  }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}

interface ReportImages {
  chartImage: string
  tableImage: string
}
```

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vidéo de formation : extraire et envoyer des images par courrier électronique à un graphique et un tableau

[![Regardez une vidéo pas à pas sur l'extraction et l'envoi par courrier électronique d'images de graphique et de tableau](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vidéo pas à pas sur l'extraction et l'envoi par courrier électronique d'images de graphique et de tableau")
