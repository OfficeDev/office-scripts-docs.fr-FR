---
title: 'Office Exemple de scénario de scripts : calculatrice de notes'
description: Exemple qui détermine le pourcentage et les notes de lettre d’une classe d’étudiants.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 2d98e68f37418ade238a707cb74cc7ccf47e8f59
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313791"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Office Exemple de scénario de scripts : calculatrice de notes

Dans ce scénario, vous êtes un instructeur qui compte les notes de fin de terme de chaque étudiant. You’ve been entering the scores for their assignments and tests as you go. À présent, il est temps de déterminer les écoles des étudiants.

Vous allez développer un script qui totale les notes pour chaque catégorie de points. Il affecte ensuite une note de lettre à chaque étudiant en fonction du total. Pour garantir la précision, vous allez ajouter quelques vérifications pour voir si des scores individuels sont trop faibles ou élevés. Si le score d’un étudiant est inférieur à zéro ou supérieur à la valeur de point possible, le script marquera la cellule avec un remplissage rouge et non le total des points de cet étudiant. Il s’agit d’une indication claire des enregistrements que vous devez vérifier. Vous allez également ajouter une mise en forme de base aux notes afin de pouvoir afficher rapidement le haut et le bas du cours.

## <a name="scripting-skills-covered"></a>Compétences d’écriture de scripts couvertes

- Mise en forme des cellules
- Vérification des erreurs
- Expressions régulières
- Mise en forme conditionnelle

## <a name="setup-instructions"></a>Instructions d’installation

1. Téléchargez <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> sur votre OneDrive.

1. Ouvrez le Excel sur le Web.

1. Sous **l’onglet Automatiser,** sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.

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
      const assignmentMaxMatches = (studentData[0][1] as string).match(/\d+/);
      const midtermMaxMatches = (studentData[0][2] as string).match(/\d+/);
      const finalMaxMatches = (studentData[0][3] as string).match(/\d+/);

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
        const total = (studentData[i][1] as number) + (studentData[i][2] as number) + (studentData[i][3] as number);
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
      let conditionalFormatting: ExcelScript.ConditionalFormat;
      conditionalFormatting = range.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
      conditionalFormatting.getCellValue().getFormat().getFont().setColor(fontColor);
      conditionalFormatting.getCellValue().getFormat().getFill().setColor(fillColor);
      conditionalFormatting.getCellValue().setRule({ formula1, operator });
    }
    ```

1. Renommez le script en calculateur **de notes** et enregistrez-le.

## <a name="running-the-script"></a>Exécution du script

Exécutez le script **Calculatrice de** notes sur la seule feuille de calcul. Le script totalera les notes et attribuera à chaque étudiant une note de lettre. Si des notes individuelles ont plus de points que la valeur de l’affectation ou du test, la note incriminée est marquée en rouge et le total n’est pas calculé. En outre, les notes « A » sont mises en surbrillant en vert, tandis que les notes « D » et « F » sont en jaune.

### <a name="before-running-the-script"></a>Avant d’exécution du script

:::image type="content" source="../../images/scenario-grade-calculator-before.png" alt-text="Feuille de calcul qui affiche des lignes de scores pour les étudiants.":::

### <a name="after-running-the-script"></a>Après l’exécution du script

:::image type="content" source="../../images/scenario-grade-calculator-after.png" alt-text="Feuille de calcul qui affiche les données de score de l’étudiant avec des cellules non valides dans des totaux rouges pour les lignes d’étudiants valides.":::
