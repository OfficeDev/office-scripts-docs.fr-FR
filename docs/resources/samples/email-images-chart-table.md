---
title: Envoyer par courrier électronique les images d Excel graphique et d'un tableau
description: Découvrez comment utiliser Office scripts et Power Automate pour extraire et envoyer par e-mail les images d'un Excel graphique et d'un tableau.
ms.date: 04/05/2021
localization_priority: Normal
ms.openlocfilehash: 0265250f7fd885cb4899d0b9493b4285496965ff
ms.sourcegitcommit: 1f003c9924e651600c913d84094506125f1055ab
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/26/2021
ms.locfileid: "52026866"
---
# <a name="use-office-scripts-and-power-automate-to-email-images-of-a-chart-and-table"></a>Utiliser Office scripts et Power Automate pour envoyer des images électroniques d'un graphique et d'un tableau

Cet exemple utilise Office scripts et Power Automate pour créer un graphique. Il envoie ensuite des images du graphique et de sa table de base par courrier électronique.

## <a name="example-scenario"></a>Exemple de scénario

* Calculer pour obtenir les derniers résultats.
* Créez un graphique.
* Obtenir des images de graphique et de tableau.
* Envoyez un e-mail aux images Power Automate.

_Données d'entrée_

:::image type="content" source="../../images/input-data.png" alt-text="Feuille de calcul montrant une table des données d'entrée.":::

_Graphique de sortie_

:::image type="content" source="../../images/chart-created.png" alt-text="Graphique en colonnes créé montrant le montant dû par le client.":::

_Courrier électronique reçu par le biais Power Automate flux_

:::image type="content" source="../../images/email-received.png" alt-text="Courrier électronique envoyé par le flux montrant le Excel graphique incorporé dans le corps.":::

## <a name="solution"></a>Solution

Cette solution est en deux parties :

1. [Un script Office pour calculer et extraire Excel graphique et tableau](#sample-code-calculate-and-extract-excel-chart-and-table)
1. Un flux Power Automate pour appeler le script et envoyer par courrier électronique les résultats. Pour obtenir un exemple sur la procédure à suivre, voir Créer un flux de travail automatisé [avec Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).

## <a name="sample-code-calculate-and-extract-excel-chart-and-table"></a>Exemple de code : calculer et extraire Excel graphique et tableau

Le script suivant calcule et extrait un Excel graphique et un tableau.

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

## <a name="power-automate-flow-email-the-chart-and-table-images"></a>Power Automate flux : envoyer un e-mail aux images du graphique et du tableau

Ce flux exécute le script et envoie par e-mail les images renvoyées.

1. Créez un **flux de cloud instantané.**
1. Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**
1. Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** avec l'action Exécuter le script (prévisualisation).  Utilisez les valeurs suivantes pour l'action :
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier**: votre classeur [(sélectionné avec le sélecateur de fichiers)](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control)
    * **Script**: nom de votre script

    :::image type="content" source="../../images/email-chart-sample-flow-1.png" alt-text="Le connecteur Excel Online (Entreprise) dans Power Automate.":::
1. Cet exemple utilise Outlook client de messagerie. Vous pouvez utiliser n'importe quel connecteur de messagerie Power Automate prend en charge, mais le reste des étapes suppose que vous avez choisi Outlook. Ajoutez **une nouvelle étape** qui utilise le connecteur **Office 365 Outlook** et l'action Envoyer et e-mail **(V2).** Utilisez les valeurs suivantes pour l'action :
    * **À**: Votre compte de messagerie de test (ou e-mail personnel)
    * **Objet :** Veuillez consulter les données du rapport
    * Pour le **champ Corps,** sélectionnez « Affichage de code » ( `</>` ), puis entrez les entrées suivantes :

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

    :::image type="content" source="../../images/email-chart-sample-flow-2.png" alt-text="Connecteur de Office 365 Outlook terminé dans Power Automate.":::
1. Enregistrez le flux et testez-le.

## <a name="training-video-extract-and-email-images-of-chart-and-table"></a>Vidéo de formation : extraire et envoyer des images par courrier électronique à un graphique et un tableau

[![Regardez une vidéo pas à pas sur l'extraction et l'envoi par courrier électronique d'images de graphique et de tableau](../../images/charts-image-vid.jpg)](https://youtu.be/152GJyqc-Kw "Vidéo pas à pas sur l'extraction et l'envoi par courrier électronique d'images de graphique et de tableau")
