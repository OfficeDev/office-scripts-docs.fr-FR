---
title: 'Exemple de scénario de scripts Office : calcul de la note'
description: Exemple qui détermine le pourcentage et les notes de lettres d’une classe d’étudiants.
ms.date: 07/24/2020
localization_priority: Normal
ms.openlocfilehash: 4e488c6cc67bda9122b88c55070654632d9c7fa2
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616739"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="0ea65-103">Exemple de scénario de scripts Office : calcul de la note</span><span class="sxs-lookup"><span data-stu-id="0ea65-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="0ea65-104">Dans ce scénario, vous êtes un instructeur qui traite les notes de fin de contrat de chaque étudiant.</span><span class="sxs-lookup"><span data-stu-id="0ea65-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="0ea65-105">Vous avez entré les scores pour les devoirs et les tests au fur et à mesure.</span><span class="sxs-lookup"><span data-stu-id="0ea65-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="0ea65-106">À présent, il est temps de déterminer les sorts des étudiants.</span><span class="sxs-lookup"><span data-stu-id="0ea65-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="0ea65-107">Vous développerez un script qui totalise les notes pour chaque catégorie de point.</span><span class="sxs-lookup"><span data-stu-id="0ea65-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="0ea65-108">Elle affecte ensuite une note à chaque étudiant en fonction du total.</span><span class="sxs-lookup"><span data-stu-id="0ea65-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="0ea65-109">Pour garantir la précision, vous allez ajouter des vérifications pour voir si des scores individuels sont trop bas ou très élevés.</span><span class="sxs-lookup"><span data-stu-id="0ea65-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="0ea65-110">Si le score d’un étudiant est inférieur à zéro ou supérieur à la valeur possible, le script marque la cellule avec un remplissage rouge et ne totalise pas les points de l’étudiant.</span><span class="sxs-lookup"><span data-stu-id="0ea65-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="0ea65-111">Il s’agit d’une indication claire des enregistrements à vérifier.</span><span class="sxs-lookup"><span data-stu-id="0ea65-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="0ea65-112">Vous ajouterez également une mise en forme de base aux notes afin de pouvoir rapidement visualiser le haut et le bas de la classe.</span><span class="sxs-lookup"><span data-stu-id="0ea65-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="0ea65-113">Compétences en matière de script</span><span class="sxs-lookup"><span data-stu-id="0ea65-113">Scripting skills covered</span></span>

- <span data-ttu-id="0ea65-114">Mise en forme de cellule</span><span class="sxs-lookup"><span data-stu-id="0ea65-114">Cell formatting</span></span>
- <span data-ttu-id="0ea65-115">Vérification des erreurs</span><span class="sxs-lookup"><span data-stu-id="0ea65-115">Error checking</span></span>
- <span data-ttu-id="0ea65-116">Expressions régulières</span><span class="sxs-lookup"><span data-stu-id="0ea65-116">Regular expressions</span></span>
- <span data-ttu-id="0ea65-117">Mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="0ea65-117">Conditional formatting</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="0ea65-118">Instructions de configuration</span><span class="sxs-lookup"><span data-stu-id="0ea65-118">Setup instructions</span></span>

1. <span data-ttu-id="0ea65-119">Téléchargez <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> vers votre espace OneDrive.</span><span class="sxs-lookup"><span data-stu-id="0ea65-119">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="0ea65-120">Ouvrez le classeur avec Excel pour le Web.</span><span class="sxs-lookup"><span data-stu-id="0ea65-120">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="0ea65-121">Sous l’onglet **automatiser** , ouvrez l' **éditeur de code**.</span><span class="sxs-lookup"><span data-stu-id="0ea65-121">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="0ea65-122">Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="0ea65-122">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the worksheet and validate the data.
      let studentsRange = workbook.getActiveWorksheet().getUsedRange();
      if (studentsRange.getColumnCount() !== 6) {
        throw new Error(`The required columns are not present. Expected column headers: "Student ID | Assignment score | Mid-term | Final | Total | Grade"`);
      }

      let studentData = studentsRange.getValues();

      // Clear the total and grade columns.
      studentsRange.getColumn(4).getCell(1, 0).getAbsoluteResizedRange(studentData.length - 1, 2).clear();

      // Clear all conditional formatting.
      workbook.getActiveWorksheet().getUsedRange().clearAllConditionalFormats();

      // Use regular expressions to read the max score from the assignment, mid-term, and final scores columns.
      let maxScores: string[] = [];
      const assignmentMaxMatches = studentData[0][1].match(/\d+/);
      const midtermMaxMatches = studentData[0][2].match(/\d+/);
      const finalMaxMatches = studentData[0][3].match(/\d+/);

      // Check the matches happened before proceeding.
      if (!(assignmentMaxMatches && midtermMaxMatches && finalMaxMatches)) {
        throw new Error(`The scores are not present in the column headers. Expected format: "Assignments (n)|Mid-term (n)|Final (n)"`);
      }

      // Use the first (and only) match from the regular expressions as the max scores.
      maxScores = [assignmentMaxMatches[0], midtermMaxMatches[0], finalMaxMatches[0]];

      // Set conditional formatting for each of the assignment, mid-term, and final scores columns.
      maxScores.forEach((score, i) => {
        let range = studentsRange.getColumn(i + 1).getCell(0, 0).getRowsBelow(studentData.length - 1);
        setCellValueConditionalFormatting(
          score,
          range,
          "#9C0006",
          "#FFC7CE",
          ExcelScript.ConditionalCellValueOperator.greaterThan
        )
      });

      // Store the current range information to avoid calling the workbook in the loop.
      let studentsRangeFormulas = studentsRange.getColumn(4).getFormulasR1C1();
      let studentsRangeValues = studentsRange.getColumn(5).getValues();

      /* Iterate over each of the student rows and compute the total score and letter grade.
      * Note that iterator starts at index 1 to skip first (header) row.
      */
      for (let i = 1; i < studentData.length; i++) {
        // If any of the scores are invalid, skip processing it.
        if (studentData[i][1] > maxScores[0] ||
          studentData[i][2] > maxScores[1] ||
          studentData[i][3] > maxScores[2]) {
          continue;
        }
        const total = studentData[i][1] + studentData[i][2] + studentData[i][3];
        let grade: string;
        switch (true) {
          case total < 60:
            grade = "F";
            break;
          case total < 70:
            grade = "D";
            break;
          case total < 80:
            grade = "C";
            break;
          case total < 90:
            grade = "B";
            break;
          default:
            grade = "A";
            break;
        }

        // Set total score formula.
        studentsRangeFormulas[i][0] = '=RC[-2]+RC[-1]';
        // Set grade cell.
        studentsRangeValues[i][0] = grade;
      }

      // Set the formulas and values outside the loop.
      studentsRange.getColumn(4).setFormulasR1C1(studentsRangeFormulas);
      studentsRange.getColumn(5).setValues(studentsRangeValues);

      // Put a conditional formatting on the grade column.
      let totalRange = studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentData.length - 1);
      setCellValueConditionalFormatting(
        "A",
        totalRange,
        "#001600",
        "#C6EFCE",
        ExcelScript.ConditionalCellValueOperator.equalTo
      );
      ["D", "F"].forEach((grade) => {
        setCellValueConditionalFormatting(
          grade,
          totalRange,
          "#443300",
          "#FFEE22",
          ExcelScript.ConditionalCellValueOperator.equalTo
        );
      })
      // Center the grade column.
      studentsRange.getColumn(5).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    }

    /**
     * Helper function to apply conditional formatting.
     * @param value Cell value to use in conditional formatting formula1.
     * @param range Target range.
     * @param fontColor Font color to use.
     * @param fillColor Fill color to use.
     * @param operator Operator to use in conditional formatting.
     */
    function setCellValueConditionalFormatting(
      value: string,
      range: ExcelScript.Range,
      fontColor: string,
      fillColor: string,
      operator: ExcelScript.ConditionalCellValueOperator) {
      // Determine the formula1 based on the type of value parameter.
      let formula1: string;
      if (isNaN(Number(value))) {
        // For cell value equalTo rule, use this format: formula1: "=\"A\"",
        formula1 = `=\"${value}\"`;
      } else {
        // For number input (greater-than or less-than rules), just append '='.
        formula1 = `=${value}`;
      }

      // Apply conditional formatting.
      let conditionalFormatting : ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({formula1, operator});
    }
    ```

5. <span data-ttu-id="0ea65-123">Renommez le script **et enregistrez** -le.</span><span class="sxs-lookup"><span data-stu-id="0ea65-123">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="0ea65-124">Exécution du script</span><span class="sxs-lookup"><span data-stu-id="0ea65-124">Running the script</span></span>

<span data-ttu-id="0ea65-125">Exécutez le script de **calculatrice de note** sur la seule feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="0ea65-125">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="0ea65-126">Le script totalise les notes et affecte à chaque étudiant une note.</span><span class="sxs-lookup"><span data-stu-id="0ea65-126">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="0ea65-127">Si des niveaux individuels ont plus de points que le devoir ou le test ne vaut, la qualité incriminée est indiquée en rouge et le total n’est pas calculé.</span><span class="sxs-lookup"><span data-stu-id="0ea65-127">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span> <span data-ttu-id="0ea65-128">De plus, toutes les notes « A » sont mises en surbrillance en vert, tandis que les notes « d » et « F » sont mises en surbrillance en jaune.</span><span class="sxs-lookup"><span data-stu-id="0ea65-128">Also, any 'A' grades are highlighted in green, while 'D' and 'F' grades are highlighted in yellow.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="0ea65-129">Avant d’exécuter le script</span><span class="sxs-lookup"><span data-stu-id="0ea65-129">Before running the script</span></span>

![Feuille de calcul qui affiche des lignes de score pour les étudiants.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="0ea65-131">Après avoir exécuté le script</span><span class="sxs-lookup"><span data-stu-id="0ea65-131">After running the script</span></span>

![Feuille de calcul qui affiche les données de score des étudiants avec des cellules non valides en rouge pour les lignes d’étudiant valides.](../../images/scenario-grade-calculator-after.png)
