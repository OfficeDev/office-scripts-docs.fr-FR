---
title: "Exemple de scénario de scripts Office : graphique des données de niveau d'eau de NOAA"
description: Exemple qui extrait des données JSON d'une base de données NOAA et les utilise pour créer un graphique.
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: ba4836cd0782ab7f2158aeaaa562c851927b90f7
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755118"
---
# <a name="office-scripts-sample-scenario-fetch-and-graph-water-level-data-from-noaa"></a>Exemple de scénario de scripts Office : extraire et graphiquer des données au niveau de l'eau à partir de la NOAA

Dans ce scénario, vous devez tracer le niveau d'eau au niveau de la station de seattle de [l'Administration nationale de l'état de Seattle.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130) Vous utiliserez des données externes pour remplir une feuille de calcul et créer un graphique.

Vous allez développer un script qui utilise la commande pour interroger la base de données `fetch` [NOAA Des descendants et des bases de données actuelles.](https://tidesandcurrents.noaa.gov/) Cela permettra d'enregistrer le niveau d'eau sur une période donnée. Les informations seront renvoyées en tant que JSON, de sorte qu'une partie du script les traduirea en valeurs de plage. Une fois que les données se trouve dans la feuille de calcul, elles sont utilisées pour créer un graphique.

## <a name="scripting-skills-covered"></a>Compétences d'écriture de scripts couvertes

- Appels d'API externes ( `fetch` )
- L'insération JSON
- Graphiques

## <a name="setup-instructions"></a>Instructions d'installation

1. Ouvrez le manuel avec Excel sur le web.

1. Sous **l'onglet Automatiser,** **sélectionnez Tous les scripts.**

1. Dans le **volet Des tâches de l'Éditeur** de code, sélectionnez **Nouveau script** et collez le script suivant dans l'éditeur.

    ```TypeScript
    /**
     * Gets data from the National Oceanic and Atmospheric Administration's Tides and Currents database. 
     * That data is used to make a chart.
     */
    async function main(workbook: ExcelScript.Workbook): Promise<void> {
      // Get the current sheet.
      let currentSheet = workbook.getActiveWorksheet();
    
      // Create selection of parameters for the fetch URL.
      // More information on the NOAA APIs is found here: 
      // https://api.tidesandcurrents.noaa.gov/api/prod/
      const option = "water_level";
      const startDate = "20201225"; /* yyyymmdd date format */
      const endDate = "20201227";
      const station = "9447130"; /* Seattle */
    
      // Construct the URL for the fetch call.
      const strQuery = `https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?product=${option}&begin_date=${startDate}&end_date=${endDate}&datum=MLLW&station=${station}&units=english&time_zone=gmt&application=NOS.COOPS.TAC.WL&format=json`;
    
      console.log(strQuery);
    
      // Resolve the Promises returned by the fetch operation.
      const response = await fetch(strQuery);
      const rawJson = await response.json();
    
      // Translate the raw JSON into a usable state.
      const stringifiedJson = JSON.stringify(rawJson);
      const noaaData = JSON.parse(stringifiedJson);
    
      // Create table headers and format them to stand out.
      let headers = [["Time", "Level"]];
      let headerRange = currentSheet.getRange("A1:B1");
      headerRange.setValues(headers);
      headerRange.getFormat().getFill().setColor("#4472C4");
      headerRange.getFormat().getFont().setColor("white");
    
      // Insert all the data in rows from JSON.
      let noaaDataCount = noaaData.data.length;
      let dataToEnter = [[], []]
      for (let i = 0; i < noaaDataCount; i++) {
        let currentDataPiece = noaaData.data[i];
        dataToEnter[i] = [currentDataPiece.t, currentDataPiece.v];
      }
    
      let dataRange = currentSheet.getRange("A2:B" + String(noaaDataCount + 1)); /* +1 to account for the title row */
      dataRange.setValues(dataToEnter);
      
      // Format the "Time" column for timestamps.
      dataRange.getColumn(0).setNumberFormatLocal("[$-en-US]mm/dd/yyyy hh:mm AM/PM;@");
    
      // Create and format a chart with the level data.
      let chart = currentSheet.addChart(ExcelScript.ChartType.xyscatterSmooth,dataRange);
      chart.getTitle().setText("Water Level - Seattle");
      chart.setTop(0);
      chart.setLeft(300);
      chart.setWidth(500);
      chart.setHeight(300);
      chart.getAxes().getValueAxis().setShowDisplayUnitLabel(false);
      chart.getAxes().getCategoryAxis().setTextOrientation(60);
      chart.getLegend().setVisible(false);

      // Add a comment with the data attribution.
      currentSheet.addComment(
        "A1", 
        `This data was taken from the National Oceanic and Atmospheric Administration's Tides and Currents database on ${new Date(Date.now())}.`
      );
    }
    ```

1. Renommez le script en **NOAA Water Level Chart** et enregistrez-le.

## <a name="running-the-script"></a>Exécution du script

Sur n'importe quelle feuille de calcul, exécutez le script **NOAA Water Level Chart.** Le script récupère les données de niveau d'eau du 25 décembre 2020 au 27 décembre 2020. Les variables au début du script peuvent être modifiées pour utiliser différentes `const` dates ou obtenir des informations de station différentes. [L'API CO-OPS pour la](https://api.tidesandcurrents.noaa.gov/api/prod/) récupération des données décrit comment obtenir toutes ces données.

### <a name="after-running-the-script"></a>Après l'exécution du script

:::image type="content" source="../../images/scenario-noaa-water-level-after.png" alt-text="La feuille de calcul après l'exécution du script affiche des données de niveau d'eau et un graphique.":::
