---
title: 'Office Exemple de scénario de scripts : analyser les téléchargements web'
description: Exemple qui prend des données de trafic Internet brutes dans un Excel et détermine l’emplacement d’origine, avant d’organiser ces informations dans une table.
ms.date: 04/27/2021
localization_priority: Normal
ms.openlocfilehash: 6c5958e9957ca49c370ae34456236bdd15f41c44
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232710"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Office Exemple de scénario de scripts : analyser les téléchargements web

Dans ce scénario, vous êtes chargé d’analyser les rapports de téléchargement à partir du site web de votre entreprise. L’objectif de cette analyse est de déterminer si le trafic web vient des États-Unis ou d’autres pays du monde.

Vos collègues téléchargent les données brutes dans votre workbook. Chaque ensemble de données de chaque semaine possède sa propre feuille de calcul. Il existe également la **feuille de calcul Résumé** avec un tableau et un graphique qui indiquent les tendances d’une semaine à l’autre.

Vous allez développer un script qui analyse les données de téléchargement hebdomadaire dans la feuille de calcul active. Elle permet d’évaluer l’adresse IP associée à chaque téléchargement et de déterminer si elle provenait ou non des États-Unis. La réponse est insérée dans la feuille de calcul sous la forme d’une valeur booléle (« TRUE » ou « FALSE ») et une mise en forme conditionnelle est appliquée à ces cellules. Les résultats de l’emplacement des adresses IP seront totaux dans la feuille de calcul et copiés dans le tableau récapitulatif.

## <a name="scripting-skills-covered"></a>Compétences d’écriture de scripts couvertes

- L’l ment de texte
- Sous-fonctions dans les scripts
- Mise en forme conditionnelle
- Tables

## <a name="setup-instructions"></a>Instructions d’installation

1. Téléchargez <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> sur votre OneDrive.

2. Ouvrez le Excel sur le Web.

3. Sous **l’onglet Automatiser,** ouvrez **Tous les scripts.**

4. Dans le **volet Des tâches de** l’Éditeur de code, appuyez sur Nouveau **script** et collez le script suivant dans l’éditeur.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      /* Get the Summary worksheet and table.
        * End the script early if either object is not in the workbook.
        */
      let summaryWorksheet = workbook.getWorksheet("Summary");
      if (!summaryWorksheet) {
        console.log("The script expects a worksheet named \"Summary\". Please download the correct template and try again.");
        return;
      }
      let summaryTable = summaryWorksheet.getTable("Table1");
      if (!summaryTable) {
        console.log("The script expects a summary table named \"Table1\". Please download the correct template and try again.");
        return;
      }
  
      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (currentWorksheet.getName().toLocaleLowerCase().indexOf("week") !== 0) {
        console.log("Please switch worksheet to one of the weekly data sheets and try again.")
        return;
      }
  
      // Get the values of the active range of the active worksheet.
      let logRange = currentWorksheet.getUsedRange();
  
      if (logRange.getColumnCount() !== 8) {
        console.log(`Verify that you are on the correct worksheet. Either the week's data has been already processed or the content is incorrect. The following columns are expected: ${[
            "Time Stamp", "IP Address", "kilobytes", "user agent code", "milliseconds", "Request", "Results", "Referrer"
        ]}`);
        return;
      }
      // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
      let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1);
  
      // Get the values of all the US IP addresses.
      let ipRange = workbook.getWorksheet("USIPAddresses").getUsedRange();
      let ipRangeValues = ipRange.getValues() as number[][];
      let logRangeValues = logRange.getValues() as string[][];
      // Remove the first row.
      let topRow = logRangeValues.shift();
      console.log(`Analyzing ${logRangeValues.length} entries.`);
  
      // Create a new array to contain the boolean representing if this is a US IP address.
      let newCol = [];
  
      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRangeValues.length; i++) {
        let curRowIP = logRangeValues[i][1];
        if (findIP(ipRangeValues, ipAddressToInteger(curRowIP)) > 0) {
          newCol.push([true]);
        } else {
          newCol.push([false]);
        }
      }
  
      // Remove the empty column header and add proper heading.
      newCol = [["Is US IP"], ...newCol];
  
      // Write the result to the spreadsheet.
      console.log(`Adding column to indicate whether IP belongs to US region or not at address: ${isUSColumn.getAddress()}`);
      console.log(newCol.length);
      console.log(newCol);
      isUSColumn.setValues(newCol);
  
      // Call the local function to add summary data to the worksheet.
      addSummaryData();
  
      // Call the local function to apply conditional formatting.
      applyConditionalFormatting(isUSColumn);
  
      // Autofit columns.
      currentWorksheet.getUsedRange().getFormat().autofitColumns();
  
      // Get the calculated summary data.
      let summaryRangeValues = currentWorksheet.getRange("J2:M2").getValues();
  
      // Add the corresponding row to the summary table.
      summaryTable.addRow(null, summaryRangeValues[0]);
      console.log("Complete.");
      return;
  
      /**
       * A function to add summary data on the worksheet.
        */
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
            "=COUNTIF(" + isUSColumn.getAddress() + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = currentWorksheet.getRange("J1:M1");
        let summaryContentRow = currentWorksheet.getRange("J2:M2");
        console.log("2");

        summaryHeaderRow.setValues(summaryHeader);
        console.log("3");

        summaryContentRow.setValues(summaryContent);
        console.log("4");

        let formats = [[".000", ".000"]];
        summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).setNumberFormats(formats);
      }
    }
    /**
     * Apply conditional formatting based on TRUE/FALSE values of the Is US IP column.
     */
    function applyConditionalFormatting(isUSColumn: ExcelScript.Range) {
      // Add conditional formatting to the new column.
      let conditionalFormatTrue = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      let conditionalFormatFalse = isUSColumn.addConditionalFormat(
          ExcelScript.ConditionalFormatType.cellValue
      );
      // Set TRUE to light blue and FALSE to light orange.
      conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#8FA8DB");
      conditionalFormatTrue.getCellValue().setRule({
          formula1: "=TRUE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
      conditionalFormatFalse.getCellValue().getFormat().getFill().setColor("#F8CCAD");
      conditionalFormatFalse.getCellValue().setRule({
          formula1: "=FALSE",
          operator: ExcelScript.ConditionalCellValueOperator.equalTo
      });
    }
    /**
     * Translate an IP address into an integer.
     * @param ipAddress: IP address to verify.
     */
    function ipAddressToInteger(ipAddress: string): number {
      // Split the IP address into octets.
      let octets = ipAddress.split(".");
  
      // Create a number for each octet and do the math to create the integer value of the IP address.
      let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
      return fullNum;
    }
    /**
     * Return the row number where the ip address is found.
     * @param ipLookupTable IP look-up table.
     * @param n IP address to number value.  
     */
    function findIP(ipLookupTable: number[][], n: number): number {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

5. Renommez le script pour **analyser les téléchargements Web** et enregistrez-le.

## <a name="running-the-script"></a>Exécution du script

Accédez à l’une **des feuilles de \* \*** calcul Semaine et exécutez le script **Analyser les téléchargements web.** Le script appliquera la mise en forme conditionnelle et la localisation sur la feuille actuelle. Il met également à jour la **feuille de calcul** Résumé.

### <a name="before-running-the-script"></a>Avant d’exécution du script

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="Feuille de calcul qui affiche les données brutes du trafic web":::

### <a name="after-running-the-script"></a>Après l’exécution du script

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="Feuille de calcul qui affiche des informations d’emplacement IP formatées avec les lignes de trafic web précédentes":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="Tableau récapitulatif et graphique récapitulant les feuilles de calcul sur lesquelles le script a été exécuté":::
