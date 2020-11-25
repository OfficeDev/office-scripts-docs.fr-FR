---
title: 'Exemple de scénario de scripts Office : analyser les téléchargements Web'
description: Exemple qui prend des données de trafic Internet brutes dans un classeur Excel et détermine l’emplacement d’origine, avant d’organiser ces informations dans un tableau.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: adc2cb401830b66b245c0dfcc4441b7ac9c8c61f
ms.sourcegitcommit: 009935c5773761c5833e5857491af47e2c95d851
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/25/2020
ms.locfileid: "49408965"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="c850b-103">Exemple de scénario de scripts Office : analyser les téléchargements Web</span><span class="sxs-lookup"><span data-stu-id="c850b-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="c850b-104">Dans ce scénario, vous êtes chargé d’analyser les rapports de téléchargement à partir du site Web de votre entreprise.</span><span class="sxs-lookup"><span data-stu-id="c850b-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="c850b-105">L’objectif de cette analyse est de déterminer si le trafic Web est en provenance des États-Unis ou ailleurs dans le monde entier.</span><span class="sxs-lookup"><span data-stu-id="c850b-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="c850b-106">Vos collègues téléchargent les données brutes dans votre classeur.</span><span class="sxs-lookup"><span data-stu-id="c850b-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="c850b-107">Chaque ensemble de données de la semaine dispose de sa propre feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c850b-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="c850b-108">Il existe également une feuille de calcul de **synthèse** contenant un tableau et un graphique présentant les tendances de semaine sur semaine.</span><span class="sxs-lookup"><span data-stu-id="c850b-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="c850b-109">Vous développerez un script qui analyse les données de téléchargements hebdomadaires dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="c850b-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="c850b-110">Elle analyse l’adresse IP associée à chaque téléchargement et détermine si elle provient ou non des États-Unis.</span><span class="sxs-lookup"><span data-stu-id="c850b-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="c850b-111">La réponse est insérée dans la feuille de calcul en tant que valeur booléenne ("TRUE" ou "FALSe") et la mise en forme conditionnelle est appliquée à ces cellules.</span><span class="sxs-lookup"><span data-stu-id="c850b-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="c850b-112">Les résultats de l’adresse IP seront totalisés sur la feuille de calcul et copiés dans le tableau récapitulatif.</span><span class="sxs-lookup"><span data-stu-id="c850b-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="c850b-113">Compétences en matière de script</span><span class="sxs-lookup"><span data-stu-id="c850b-113">Scripting skills covered</span></span>

- <span data-ttu-id="c850b-114">Analyse de texte</span><span class="sxs-lookup"><span data-stu-id="c850b-114">Text parsing</span></span>
- <span data-ttu-id="c850b-115">Sous-fonctions dans les scripts</span><span class="sxs-lookup"><span data-stu-id="c850b-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="c850b-116">Mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="c850b-116">Conditional formatting</span></span>
- <span data-ttu-id="c850b-117">Tables</span><span class="sxs-lookup"><span data-stu-id="c850b-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="c850b-118">Vidéo de démonstration</span><span class="sxs-lookup"><span data-stu-id="c850b-118">Demo video</span></span>

<span data-ttu-id="c850b-119">Cet exemple a été démo dans le cadre de l’appel de la communauté de développeurs de compléments Office pour le 2020 février.</span><span class="sxs-lookup"><span data-stu-id="c850b-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

> [!NOTE]
> <span data-ttu-id="c850b-120">Le code présenté dans cette vidéo utilise un modèle d’API plus ancien ( [API Async pour les scripts Office](../../develop/excel-async-model.md)).</span><span class="sxs-lookup"><span data-stu-id="c850b-120">The code shown in this video uses an older API model (the [Office Scripts Async APIs](../../develop/excel-async-model.md)).</span></span> <span data-ttu-id="c850b-121">L’exemple présenté sur cette page a été mis à jour, mais le code semble un peu différent de l’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="c850b-121">The sample presented on this page has been updated, but the code looks a little different from the recording.</span></span> <span data-ttu-id="c850b-122">Les modifications n’affectent pas le comportement du script ou de l’autre contenu dans la démonstration du présentateur.</span><span class="sxs-lookup"><span data-stu-id="c850b-122">The changes don't affect the behavior of the script or the other content in the presenter's demo.</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="c850b-123">Instructions de configuration</span><span class="sxs-lookup"><span data-stu-id="c850b-123">Setup instructions</span></span>

1. <span data-ttu-id="c850b-124">Téléchargez <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> vers votre espace OneDrive.</span><span class="sxs-lookup"><span data-stu-id="c850b-124">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="c850b-125">Ouvrez le classeur avec Excel pour le Web.</span><span class="sxs-lookup"><span data-stu-id="c850b-125">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="c850b-126">Sous l’onglet **automatiser** , ouvrez l' **éditeur de code**.</span><span class="sxs-lookup"><span data-stu-id="c850b-126">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="c850b-127">Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="c850b-127">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="c850b-128">Renommez le script pour analyser et enregistrer les **téléchargements Web** .</span><span class="sxs-lookup"><span data-stu-id="c850b-128">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="c850b-129">Exécution du script</span><span class="sxs-lookup"><span data-stu-id="c850b-129">Running the script</span></span>

<span data-ttu-id="c850b-130">Accédez à l’une des feuilles de calcul de **semaine \* \*** et exécutez le script **analyze Web Downloads** .</span><span class="sxs-lookup"><span data-stu-id="c850b-130">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="c850b-131">Le script applique la mise en forme conditionnelle et l’étiquetage de l’emplacement sur la feuille actuelle.</span><span class="sxs-lookup"><span data-stu-id="c850b-131">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="c850b-132">Elle met également à jour la feuille de calcul de **synthèse** .</span><span class="sxs-lookup"><span data-stu-id="c850b-132">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="c850b-133">Avant d’exécuter le script</span><span class="sxs-lookup"><span data-stu-id="c850b-133">Before running the script</span></span>

![Feuille de calcul qui affiche les données de trafic Web brut.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="c850b-135">Après avoir exécuté le script</span><span class="sxs-lookup"><span data-stu-id="c850b-135">After running the script</span></span>

![Feuille de calcul qui montre les informations d’emplacement IP mises en forme avec les lignes de trafic Web précédentes.](../../images/scenario-analyze-web-downloads-after.png)

![Tableau récapitulatif et graphique résumant les feuilles de calcul sur lesquelles le script a été exécuté.](../../images/scenario-analyze-web-downloads-table.png)
