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
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Exemple de scénario de scripts Office : calcul de la note

Dans ce scénario, vous êtes un instructeur qui traite les notes de fin de contrat de chaque étudiant. Vous avez entré les scores pour les devoirs et les tests au fur et à mesure. À présent, il est temps de déterminer les sorts des étudiants.

Vous développerez un script qui totalise les notes pour chaque catégorie de point. Elle affecte ensuite une note à chaque étudiant en fonction du total. Pour garantir la précision, vous allez ajouter des vérifications pour voir si des scores individuels sont trop bas ou très élevés. Si le score d’un étudiant est inférieur à zéro ou supérieur à la valeur possible, le script marque la cellule avec un remplissage rouge et ne totalise pas les points de l’étudiant. Il s’agit d’une indication claire des enregistrements à vérifier. Vous ajouterez également une mise en forme de base aux notes afin de pouvoir rapidement visualiser le haut et le bas de la classe.

## <a name="scripting-skills-covered"></a>Compétences en matière de script

- Mise en forme de cellule
- Vérification des erreurs
- Expressions régulières
- Mise en forme conditionnelle

## <a name="setup-instructions"></a>Instructions de configuration

1. Téléchargez <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> vers votre espace OneDrive.

2. Ouvrez le classeur avec Excel pour le Web.

3. Sous l’onglet **automatiser** , ouvrez l' **éditeur de code**.

4. Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.

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

5. Renommez le script **et enregistrez** -le.

## <a name="running-the-script"></a>Exécution du script

Exécutez le script de **calculatrice de note** sur la seule feuille de calcul. Le script totalise les notes et affecte à chaque étudiant une note. Si des niveaux individuels ont plus de points que le devoir ou le test ne vaut, la qualité incriminée est indiquée en rouge et le total n’est pas calculé. De plus, toutes les notes « A » sont mises en surbrillance en vert, tandis que les notes « d » et « F » sont mises en surbrillance en jaune.

### <a name="before-running-the-script"></a>Avant d’exécuter le script

![Feuille de calcul qui affiche des lignes de score pour les étudiants.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>Après avoir exécuté le script

![Feuille de calcul qui affiche les données de score des étudiants avec des cellules non valides en rouge pour les lignes d’étudiant valides.](../../images/scenario-grade-calculator-after.png)
