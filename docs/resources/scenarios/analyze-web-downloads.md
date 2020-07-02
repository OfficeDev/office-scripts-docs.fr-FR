---
title: 'Exemple de scénario de scripts Office : analyser les téléchargements Web'
description: Exemple qui prend des données de trafic Internet brutes dans un classeur Excel et détermine l’emplacement d’origine, avant d’organiser ces informations dans un tableau.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 425d2af432d6b3c4b7604daf7935d2cc1ec059a8
ms.sourcegitcommit: 30750c4392db3ef057075a5702abb92863c93eda
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/01/2020
ms.locfileid: "44999266"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Exemple de scénario de scripts Office : analyser les téléchargements Web

Dans ce scénario, vous êtes chargé d’analyser les rapports de téléchargement à partir du site Web de votre entreprise. L’objectif de cette analyse est de déterminer si le trafic Web est en provenance des États-Unis ou ailleurs dans le monde entier.

Vos collègues téléchargent les données brutes dans votre classeur. Chaque ensemble de données de la semaine dispose de sa propre feuille de calcul. Il existe également une feuille de calcul de **synthèse** contenant un tableau et un graphique présentant les tendances de semaine sur semaine.

Vous développerez un script qui analyse les données de téléchargements hebdomadaires dans la feuille de calcul active. Elle analyse l’adresse IP associée à chaque téléchargement et détermine si elle provient ou non des États-Unis. La réponse est insérée dans la feuille de calcul en tant que valeur booléenne ("TRUE" ou "FALSe") et la mise en forme conditionnelle est appliquée à ces cellules. Les résultats de l’adresse IP seront totalisés sur la feuille de calcul et copiés dans le tableau récapitulatif.

## <a name="scripting-skills-covered"></a>Compétences en matière de script

- Analyse de texte
- Sous-fonctions dans les scripts
- Mise en forme conditionnelle
- Tables

## <a name="demo-video"></a>Vidéo de démonstration

Cet exemple a été démo dans le cadre de l’appel de la communauté de développeurs de compléments Office pour le 2020 février.

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

> [!NOTE]
> Le code présenté dans cette vidéo utilise un modèle d’API plus ancien ( [API Async pour les scripts Office](../../develop/excel-async-model.md)). L’exemple présenté sur cette page a été mis à jour, mais le code semble un peu différent de l’enregistrement. Les modifications n’affectent pas le comportement du script ou de l’autre contenu dans la démonstration du présentateur.

## <a name="setup-instructions"></a>Instructions de configuration

1. Téléchargez <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> vers votre espace OneDrive.

2. Ouvrez le classeur avec Excel pour le Web.

3. Sous l’onglet **automatiser** , ouvrez l' **éditeur de code**.

4. Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the Summary worksheet and table.
      let summaryWorksheet = workbook.getWorksheet("Summary");
      let summaryTable = summaryWorksheet?.getTable("Table1");
      if (!summaryWorksheet || !summaryTable) {
          console.log("The script expects the Summary worksheet with a summary table named Table1. Please download the correct template and try again.");
          return;
      }
  
      // Get the current worksheet.
      let currentWorksheet = workbook.getActiveWorksheet();
      if (!currentWorksheet.getName().toLocaleLowerCase().startsWith("week")) {
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
      let ipRangeValues = ipRange.getValues();
      let logRangeValues = logRange.getValues();
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
        let summaryHeaderRow = currentWorksheet
            .getRange("J1:M1");
        let summaryContentRow = currentWorksheet
            .getRange("J2:M2");
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
        conditionalFormatTrue.getCellValue().getFormat().getFill().setColor("#F8CCAD");
        conditionalFormatTrue.getCellValue().setRule({
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

5. Renommez le script pour analyser et enregistrer les **téléchargements Web** .

## <a name="running-the-script"></a>Exécution du script

Accédez à l’une des feuilles de calcul de **semaine \* \* ** et exécutez le script **analyze Web Downloads** . Le script applique la mise en forme conditionnelle et l’étiquetage de l’emplacement sur la feuille actuelle. Elle met également à jour la feuille de calcul de **synthèse** .

### <a name="before-running-the-script"></a>Avant d’exécuter le script

![Feuille de calcul qui affiche les données de trafic Web brut.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>Après avoir exécuté le script

![Feuille de calcul qui montre les informations d’emplacement IP mises en forme avec les lignes de trafic Web précédentes.](../../images/scenario-analyze-web-downloads-after.png)

![Tableau récapitulatif et graphique résumant les feuilles de calcul sur lesquelles le script a été exécuté.](../../images/scenario-analyze-web-downloads-table.png)
