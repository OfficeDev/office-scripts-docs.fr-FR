---
title: 'Office Exemple de scénario de scripts : analyser les téléchargements web'
description: Exemple qui prend des données de trafic Internet brutes dans un Excel et détermine l’emplacement d’origine, avant d’organiser ces informations dans une table.
ms.date: 04/27/2021
localization_priority: Normal
ms.openlocfilehash: bdd6b43290e5432d87c4a85a35fbaf32967fbf03
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074458"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="78af6-103">Office Exemple de scénario de scripts : analyser les téléchargements web</span><span class="sxs-lookup"><span data-stu-id="78af6-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="78af6-104">Dans ce scénario, vous êtes chargé d’analyser les rapports de téléchargement à partir du site web de votre entreprise.</span><span class="sxs-lookup"><span data-stu-id="78af6-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="78af6-105">L’objectif de cette analyse est de déterminer si le trafic web vient des États-Unis ou d’autres pays du monde.</span><span class="sxs-lookup"><span data-stu-id="78af6-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="78af6-106">Vos collègues téléchargent les données brutes dans votre workbook.</span><span class="sxs-lookup"><span data-stu-id="78af6-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="78af6-107">Chaque ensemble de données de chaque semaine possède sa propre feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="78af6-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="78af6-108">Il existe également la **feuille de calcul Résumé** avec un tableau et un graphique qui indiquent les tendances d’une semaine à l’autre.</span><span class="sxs-lookup"><span data-stu-id="78af6-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="78af6-109">Vous allez développer un script qui analyse les données de téléchargement hebdomadaires dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="78af6-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="78af6-110">Elle permet d’évaluer l’adresse IP associée à chaque téléchargement et de déterminer si elle provenait ou non des États-Unis.</span><span class="sxs-lookup"><span data-stu-id="78af6-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="78af6-111">La réponse est insérée dans la feuille de calcul sous la forme d’une valeur booléle (« TRUE » ou « FALSE ») et une mise en forme conditionnelle est appliquée à ces cellules.</span><span class="sxs-lookup"><span data-stu-id="78af6-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="78af6-112">Les résultats de l’emplacement des adresses IP seront totaux dans la feuille de calcul et copiés dans le tableau récapitulatif.</span><span class="sxs-lookup"><span data-stu-id="78af6-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="78af6-113">Compétences d’écriture de scripts couvertes</span><span class="sxs-lookup"><span data-stu-id="78af6-113">Scripting skills covered</span></span>

- <span data-ttu-id="78af6-114">L’l ment de texte</span><span class="sxs-lookup"><span data-stu-id="78af6-114">Text parsing</span></span>
- <span data-ttu-id="78af6-115">Sous-fonctions dans les scripts</span><span class="sxs-lookup"><span data-stu-id="78af6-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="78af6-116">Mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="78af6-116">Conditional formatting</span></span>
- <span data-ttu-id="78af6-117">Tables</span><span class="sxs-lookup"><span data-stu-id="78af6-117">Tables</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="78af6-118">Instructions d’installation</span><span class="sxs-lookup"><span data-stu-id="78af6-118">Setup instructions</span></span>

1. <span data-ttu-id="78af6-119">Téléchargez <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> sur votre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="78af6-119">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="78af6-120">Ouvrez le Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="78af6-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="78af6-121">Sous **l’onglet Automatiser,** ouvrez **Tous les scripts.**</span><span class="sxs-lookup"><span data-stu-id="78af6-121">Under the **Automate** tab, open **All Scripts**.</span></span>

4. <span data-ttu-id="78af6-122">Dans le **volet Des tâches de l’Éditeur** de code, appuyez **sur Nouveau script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="78af6-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="78af6-123">Renommez le script pour **analyser les téléchargements web** et enregistrez-le.</span><span class="sxs-lookup"><span data-stu-id="78af6-123">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="78af6-124">Exécution du script</span><span class="sxs-lookup"><span data-stu-id="78af6-124">Running the script</span></span>

<span data-ttu-id="78af6-125">Accédez à l’une **des feuilles de \* \*** calcul Semaine et exécutez le script **Analyser les téléchargements web.**</span><span class="sxs-lookup"><span data-stu-id="78af6-125">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="78af6-126">Le script appliquera la mise en forme conditionnelle et la localisation sur la feuille actuelle.</span><span class="sxs-lookup"><span data-stu-id="78af6-126">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="78af6-127">Il met également à jour la **feuille de calcul** Résumé.</span><span class="sxs-lookup"><span data-stu-id="78af6-127">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="78af6-128">Avant d’exécution du script</span><span class="sxs-lookup"><span data-stu-id="78af6-128">Before running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-before.png" alt-text="Feuille de calcul qui affiche les données brutes du trafic web.":::

### <a name="after-running-the-script"></a><span data-ttu-id="78af6-130">Après l’exécution du script</span><span class="sxs-lookup"><span data-stu-id="78af6-130">After running the script</span></span>

:::image type="content" source="../../images/scenario-analyze-web-downloads-after.png" alt-text="Feuille de calcul qui affiche des informations d’emplacement IP formatées avec les lignes de trafic web précédentes.":::

:::image type="content" source="../../images/scenario-analyze-web-downloads-table.png" alt-text="Tableau récapitulatif et graphique récapitulant les feuilles de calcul sur lesquelles le script a été exécuté.":::
