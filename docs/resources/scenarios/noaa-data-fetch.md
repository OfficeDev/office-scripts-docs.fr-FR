---
title: 'Exemple de scénario de scripts Office : graphique des données de niveau d’eau de NOAA'
description: Exemple qui extrait des données JSON d’une base de données NOAA et les utilise pour créer un graphique.
ms.date: 01/05/2021
localization_priority: Normal
ms.openlocfilehash: d2afcd05125ea66c028d8e21bcc878371c20fcc3
ms.sourcegitcommit: 30c4b731dc8d18fca5aa74ce59e18a4a63eb4ffc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784186"
---
# <a name="office-scripts-sample-scenario-graph-water-level-data-from-noaa"></a><span data-ttu-id="774cd-103">Exemple de scénario de scripts Office : graphique des données de niveau d’eau de NOAA</span><span class="sxs-lookup"><span data-stu-id="774cd-103">Office Scripts sample scenario: Graph water-level data from NOAA</span></span>

<span data-ttu-id="774cd-104">Dans ce scénario, vous devez tracer le niveau d’eau au niveau de la station De Seattle de [l’Administration nationale de l’unité nationale d’administration.](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130)</span><span class="sxs-lookup"><span data-stu-id="774cd-104">In this scenario, you need to plot the water level at the [National Oceanic and Atmospheric Administration's Seattle station](https://tidesandcurrents.noaa.gov/stationhome.html?id=9447130).</span></span> <span data-ttu-id="774cd-105">Vous utiliserez des données externes pour remplir une feuille de calcul et créer un graphique.</span><span class="sxs-lookup"><span data-stu-id="774cd-105">You'll use external data to populate a spreadsheet and create a chart.</span></span>

<span data-ttu-id="774cd-106">Vous allez développer un script qui utilise la commande pour interroger la base de données `fetch` [NOAA Des descendants et des bases de données actuelles.](https://tidesandcurrents.noaa.gov/)</span><span class="sxs-lookup"><span data-stu-id="774cd-106">You'll develop a script that uses the `fetch` command to query the [NOAA Tides and Currents database](https://tidesandcurrents.noaa.gov/).</span></span> <span data-ttu-id="774cd-107">Cela permettra d’enregistrer le niveau d’eau sur une période donnée.</span><span class="sxs-lookup"><span data-stu-id="774cd-107">That will get the water level recorded across a given time span.</span></span> <span data-ttu-id="774cd-108">Les informations seront renvoyées en tant que JSON, de sorte qu’une partie du script les traduirea en valeurs de plage.</span><span class="sxs-lookup"><span data-stu-id="774cd-108">The information will be returned as JSON, so part of the script will translate that into range values.</span></span> <span data-ttu-id="774cd-109">Une fois que les données se trouve dans la feuille de calcul, elles sont utilisées pour créer un graphique.</span><span class="sxs-lookup"><span data-stu-id="774cd-109">Once the data is in the spreadsheet, it will be used to make a chart.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="774cd-110">Compétences d’écriture de scripts couvertes</span><span class="sxs-lookup"><span data-stu-id="774cd-110">Scripting skills covered</span></span>

- <span data-ttu-id="774cd-111">Appels d’API externes ( `fetch` )</span><span class="sxs-lookup"><span data-stu-id="774cd-111">External API calls (`fetch`)</span></span>
- <span data-ttu-id="774cd-112">L’insération JSON</span><span class="sxs-lookup"><span data-stu-id="774cd-112">JSON parsing</span></span>
- <span data-ttu-id="774cd-113">Graphiques</span><span class="sxs-lookup"><span data-stu-id="774cd-113">Charts</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="774cd-114">Instructions d’installation</span><span class="sxs-lookup"><span data-stu-id="774cd-114">Setup instructions</span></span>

1. <span data-ttu-id="774cd-115">Ouvrez le manuel avec Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="774cd-115">Open the workbook with Excel on the web.</span></span>

1. <span data-ttu-id="774cd-116">Sous **l’onglet Automatiser,** **sélectionnez Tous les scripts.**</span><span class="sxs-lookup"><span data-stu-id="774cd-116">Under the **Automate** tab, select **All Scripts**.</span></span>

1. <span data-ttu-id="774cd-117">Dans le **volet Des tâches de l’Éditeur** de code, sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="774cd-117">In the **Code Editor** task pane, select **New Script** and paste the following script into the editor.</span></span>

    ```typescript
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

1. <span data-ttu-id="774cd-118">Renommez le script en **NOAA Water Level Chart** et enregistrez-le.</span><span class="sxs-lookup"><span data-stu-id="774cd-118">Rename the script to **NOAA Water Level Chart** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="774cd-119">Exécution du script</span><span class="sxs-lookup"><span data-stu-id="774cd-119">Running the script</span></span>

<span data-ttu-id="774cd-120">Sur n’importe quelle feuille de calcul, exécutez le script **NOAA Water Level Chart.**</span><span class="sxs-lookup"><span data-stu-id="774cd-120">On any worksheet, run the **NOAA Water Level Chart** script.</span></span> <span data-ttu-id="774cd-121">Le script récupère les données de niveau d’eau du 25 décembre 2020 au 27 décembre 2020.</span><span class="sxs-lookup"><span data-stu-id="774cd-121">The script fetches the water level data from December 25, 2020 to December 27, 2020.</span></span> <span data-ttu-id="774cd-122">Les variables au début du script peuvent être modifiées pour utiliser des dates différentes ou `const` obtenir des informations de station différentes.</span><span class="sxs-lookup"><span data-stu-id="774cd-122">The `const` variables at the beginning of the script can be changed to use different dates or get different station information.</span></span> <span data-ttu-id="774cd-123">[L’API CO-OPS pour](https://api.tidesandcurrents.noaa.gov/api/prod/) la récupération des données décrit comment obtenir toutes ces données.</span><span class="sxs-lookup"><span data-stu-id="774cd-123">The [CO-OPS API For Data Retrieval](https://api.tidesandcurrents.noaa.gov/api/prod/) describes how to get all this data.</span></span>

### <a name="after-running-the-script"></a><span data-ttu-id="774cd-124">Après l’exécution du script</span><span class="sxs-lookup"><span data-stu-id="774cd-124">After running the script</span></span>

![La feuille de calcul après l’exécution du script affiche des données de niveau d’eau et un graphique.](../../images/scenario-noaa-water-level-after.png)
