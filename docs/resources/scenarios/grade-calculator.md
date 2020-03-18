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
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="fc2e1-103">Exemple de scénario de scripts Office : calcul de la note</span><span class="sxs-lookup"><span data-stu-id="fc2e1-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="fc2e1-104">Dans ce scénario, vous êtes un instructeur qui traite les notes de fin de contrat de chaque étudiant.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="fc2e1-105">Vous avez entré les scores pour les devoirs et les tests au fur et à mesure.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="fc2e1-106">À présent, il est temps de déterminer les sorts des étudiants.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="fc2e1-107">Vous développerez un script qui totalise les notes pour chaque catégorie de point.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="fc2e1-108">Elle affecte ensuite une note à chaque étudiant en fonction du total.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="fc2e1-109">Pour garantir la précision, vous allez ajouter des vérifications pour voir si des scores individuels sont trop bas ou très élevés.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="fc2e1-110">Si le score d’un étudiant est inférieur à zéro ou supérieur à la valeur possible, le script marque la cellule avec un remplissage rouge et ne totalise pas les points de l’étudiant.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="fc2e1-111">Il s’agit d’une indication claire des enregistrements à vérifier.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="fc2e1-112">Vous ajouterez également une mise en forme de base aux notes afin de pouvoir rapidement visualiser le haut et le bas de la classe.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="fc2e1-113">Compétences en matière de script</span><span class="sxs-lookup"><span data-stu-id="fc2e1-113">Scripting skills covered</span></span>

- <span data-ttu-id="fc2e1-114">Mise en forme de cellule</span><span class="sxs-lookup"><span data-stu-id="fc2e1-114">Cell formatting</span></span>
- <span data-ttu-id="fc2e1-115">Vérification des erreurs</span><span class="sxs-lookup"><span data-stu-id="fc2e1-115">Error checking</span></span>
- <span data-ttu-id="fc2e1-116">Expressions régulières</span><span class="sxs-lookup"><span data-stu-id="fc2e1-116">Regular expressions</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="fc2e1-117">Instructions de configuration</span><span class="sxs-lookup"><span data-stu-id="fc2e1-117">Setup instructions</span></span>

1. <span data-ttu-id="fc2e1-118">Téléchargez <a href="grade-calculator.xlsx">grade-Calculator. xlsx</a> sur votre OneDrive.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-118">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="fc2e1-119">Ouvrez le classeur avec Excel pour le Web.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-119">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="fc2e1-120">Sous l’onglet **automatiser** , ouvrez l' **éditeur de code**.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-120">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="fc2e1-121">Dans le volet Office **éditeur de code** , appuyez sur **nouveau script** et collez le script suivant dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-121">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="fc2e1-122">Renommez le script **et enregistrez** -le.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="fc2e1-123">Exécution du script</span><span class="sxs-lookup"><span data-stu-id="fc2e1-123">Running the script</span></span>

<span data-ttu-id="fc2e1-124">Exécutez le script de **calculatrice de note** sur la seule feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="fc2e1-125">Le script totalise les notes et affecte à chaque étudiant une note.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="fc2e1-126">Si des niveaux individuels ont plus de points que le devoir ou le test ne vaut, la qualité incriminée est indiquée en rouge et le total n’est pas calculé.</span><span class="sxs-lookup"><span data-stu-id="fc2e1-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="fc2e1-127">Avant d’exécuter le script</span><span class="sxs-lookup"><span data-stu-id="fc2e1-127">Before running the script</span></span>

![Feuille de calcul qui affiche des lignes de score pour les étudiants.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="fc2e1-129">Après avoir exécuté le script</span><span class="sxs-lookup"><span data-stu-id="fc2e1-129">After running the script</span></span>

![Feuille de calcul qui affiche les données de score des étudiants avec des cellules non valides en rouge pour les lignes d’étudiant valides.](../../images/scenario-grade-calculator-after.png)
