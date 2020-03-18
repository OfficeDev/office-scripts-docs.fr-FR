---
title: 'Exemple de scénario de scripts Office : calcul de la note'
description: Exemple qui détermine le pourcentage et les notes de lettres d’une classe d’étudiants.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 0db6f7c116594f7655bfc0adc8f5a79dbbf2a0af
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700233"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Exemple de scénario de scripts Office : calcul de la note

Dans ce scénario, vous êtes un instructeur qui traite les notes de fin de contrat de chaque étudiant. Vous avez entré les scores pour les devoirs et les tests au fur et à mesure. À présent, il est temps de déterminer les sorts des étudiants.

Vous développerez un script qui totalise les notes pour chaque catégorie de point. Elle affecte ensuite une note à chaque étudiant en fonction du total. Pour garantir la précision, vous allez ajouter des vérifications pour voir si des scores individuels sont trop bas ou très élevés. Si le score d’un étudiant est inférieur à zéro ou supérieur à la valeur possible, le script marque la cellule avec un remplissage rouge et ne totalise pas les points de l’étudiant. Il s’agit d’une indication claire des enregistrements à vérifier. Vous ajouterez également une mise en forme de base aux notes afin de pouvoir rapidement visualiser le haut et le bas de la classe.

## <a name="scripting-skills-covered"></a>Compétences en matière de script

- Mise en forme de cellule
- Vérification des erreurs
- Expressions régulières

## <a name="setup-instructions"></a>Instructions de configuration

1. Téléchargez <a href="grade-calculator.xlsx">grade-Calculator. xlsx</a> sur votre OneDrive.

2. Ouvrez le classeur avec Excel pour le Web.

3. Sous l’onglet **automatiser** , ouvrez l' **éditeur de code**.

4. Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the number of student record rows.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let studentsRange = sheet.getUsedRange().load("values, rowCount");
      await context.sync();
      console.log("Total students: " + (studentsRange.rowCount - 1));

      // Clean up any formatting from previous runs of the script.
      studentsRange.clear(Excel.ClearApplyTo.formats);
      studentsRange.getColumn(4).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      await context.sync();

      // Parse the headers for the maximum possible scores for each category.
      // The format is `category (score)`.
      let assignmentsMax = studentsRange.values[0][1].match(/\d+/)[0];
      let midTermMax = studentsRange.values[0][2].match(/\d+/)[0];
      let finalsMax = studentsRange.values[0][3].match(/\d+/)[0];
      console.log("Assignments max score:" + assignmentsMax);
      console.log("Mid-term max score: " + midTermMax);
      console.log("Final max score: " + finalsMax);

      // Look at every student row.
      for (let i = 1; i < studentsRange.values.length; i++) {
        let row = studentsRange.values[i];
        let total = row[1] + row[2] + row[3];
        let valid = true;

        // Look for any records that are too low or too high.
        if (row[1] < 0 || row[1] > assignmentsMax) {
          studentsRange.getCell(i, 1).format.fill.color = "Red";
          valid = false;
        }
        if (row[2] < 0 || row[2] > midTermMax) {
          studentsRange.getCell(i, 2).format.fill.color = "Red";
          valid = false;
        }
        if (row[3] < 0 || row[3] > finalsMax) {
          studentsRange.getCell(i, 3).format.fill.color = "Red";
          valid = false;
        }

        // If the scores are valid, total that student's points and assign them a letter grade.
        if (valid) {
          let grade: string;
          switch (true) {
            case total < 60:
              grade = "E";
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

          studentsRange.getCell(i, 4).values = [[total]];
          studentsRange.getCell(i, 5).values = [[grade]];

          // Highlight excellent students and those in need of attention.
          if (grade === "A") {
            studentsRange.getCell(i, 5).format.fill.color = "Green";
          } else if (grade === "E" || grade === "D") {
            studentsRange.getCell(i, 5).format.fill.color = "Orange";
          }
        }
      }

      studentsRange.getColumn(5).format.horizontalAlignment = "Center";
    }
    ```

5. Renommez le script **et enregistrez** -le.

## <a name="running-the-script"></a>Exécution du script

Exécutez le script de **calculatrice de note** sur la seule feuille de calcul. Le script totalise les notes et affecte à chaque étudiant une note. Si des niveaux individuels ont plus de points que le devoir ou le test ne vaut, la qualité incriminée est indiquée en rouge et le total n’est pas calculé.

### <a name="before-running-the-script"></a>Avant d’exécuter le script

![Feuille de calcul qui affiche des lignes de score pour les étudiants.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>Après avoir exécuté le script

![Feuille de calcul qui affiche les données de score des étudiants avec des cellules non valides en rouge pour les lignes d’étudiant valides.](../../images/scenario-grade-calculator-after.png)
