---
title: Lire les données d’un classeur avec les scripts Office d’Excel pour le web
description: Didacticiel des scripts Office sur la lecture de données à partir de classeurs et l’évaluation de ces données dans le script.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700181"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="9229c-103">Lire les données d’un classeur avec les scripts Office d’Excel pour le web</span><span class="sxs-lookup"><span data-stu-id="9229c-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="9229c-104">Ce didacticiel vous apprend comment lire des données à partir d’un classeur à l’aide d’un script Office pour Excel pour le web.</span><span class="sxs-lookup"><span data-stu-id="9229c-104">This tutorial will teach you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="9229c-105">Vous pourrez ensuite modifier les données que vous avez lues et les replacer dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="9229c-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="9229c-106">Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="9229c-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9229c-107">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="9229c-107">Prerequisites</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

<span data-ttu-id="9229c-108">Avant de commencer ce didacticiel, vous devez disposer d’un accès aux scripts Office, ce qui nécessite ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="9229c-108">Before starting this tutorial, you'll need access to Office Scripts, which requires the following:</span></span>

- <span data-ttu-id="9229c-109">[Excel pour le web](https://www.office.com/launch/excel).</span><span class="sxs-lookup"><span data-stu-id="9229c-109">[Excel on the web](https://www.office.com/launch/excel).</span></span>
- <span data-ttu-id="9229c-110">Demandez à votre administrateur d’[activer les scripts Office pour votre organisation](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), ce qui ajoute l’onglet **Automatiser** au ruban.</span><span class="sxs-lookup"><span data-stu-id="9229c-110">Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9229c-111">Ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="9229c-111">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="9229c-112">Si vous débutez avec JavaScript, nous vous conseillons de consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="9229c-112">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="9229c-113">Rendez-vous sur [Scripts Office dans Excel pour le web](../overview/excel.md) pour en savoir plus sur l’environnement de script.</span><span class="sxs-lookup"><span data-stu-id="9229c-113">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="9229c-114">Lire une cellule</span><span class="sxs-lookup"><span data-stu-id="9229c-114">Read a cell</span></span>

<span data-ttu-id="9229c-115">Les scripts créés avec l’enregistreur d’actions peuvent uniquement écrire des informations dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="9229c-115">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="9229c-116">L’éditeur de code vous permet de modifier et de créer des scripts qui peuvent également lire les données d’un classeur.</span><span class="sxs-lookup"><span data-stu-id="9229c-116">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="9229c-117">Nous allons créer un script qui lit les données et agit en fonction de ce qui a été lu.</span><span class="sxs-lookup"><span data-stu-id="9229c-117">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="9229c-118">Nous allons utiliser un exemple de relevé bancaire.</span><span class="sxs-lookup"><span data-stu-id="9229c-118">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="9229c-119">Il s’agit d’un relevé combiné de compte courant et de crédit.</span><span class="sxs-lookup"><span data-stu-id="9229c-119">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="9229c-120">Malheureusement, les changements de soldes sont rapportés différemment.</span><span class="sxs-lookup"><span data-stu-id="9229c-120">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="9229c-121">Le relevé de compte courant donne les revenus comme crédit positif et les dépenses comme débit négatif.</span><span class="sxs-lookup"><span data-stu-id="9229c-121">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="9229c-122">Le relevé de crédit fait l’inverse.</span><span class="sxs-lookup"><span data-stu-id="9229c-122">The credit statement does the opposite.</span></span>

<span data-ttu-id="9229c-123">Dans le reste du didacticiel, nous allons normaliser ces données à l’aide d’un script.</span><span class="sxs-lookup"><span data-stu-id="9229c-123">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="9229c-124">Pour commencer, voyons comment lire des données à partir du classeur.</span><span class="sxs-lookup"><span data-stu-id="9229c-124">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="9229c-125">Créez une nouvelle feuille de calcul dans le classeur courant, vous l’utiliserez pour le reste du didacticiel.</span><span class="sxs-lookup"><span data-stu-id="9229c-125">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="9229c-126">Copiez les données suivantes et collez-les dans la feuille de calcul en commençant à la cellule **A1**.</span><span class="sxs-lookup"><span data-stu-id="9229c-126">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="9229c-127">Date</span><span class="sxs-lookup"><span data-stu-id="9229c-127">Date</span></span> |<span data-ttu-id="9229c-128">Compte</span><span class="sxs-lookup"><span data-stu-id="9229c-128">Account</span></span> |<span data-ttu-id="9229c-129">Description</span><span class="sxs-lookup"><span data-stu-id="9229c-129">Description</span></span> |<span data-ttu-id="9229c-130">Débit</span><span class="sxs-lookup"><span data-stu-id="9229c-130">Debit</span></span> |<span data-ttu-id="9229c-131">Crédit</span><span class="sxs-lookup"><span data-stu-id="9229c-131">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="9229c-132">10/10/2019</span><span class="sxs-lookup"><span data-stu-id="9229c-132">10/10/2019</span></span> |<span data-ttu-id="9229c-133">Compte courant</span><span class="sxs-lookup"><span data-stu-id="9229c-133">Checking</span></span> |<span data-ttu-id="9229c-134">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="9229c-134">Coho Vineyard</span></span> |<span data-ttu-id="9229c-135">−20,05</span><span class="sxs-lookup"><span data-stu-id="9229c-135">-20.05</span></span> | |
    |<span data-ttu-id="9229c-136">11/10/2019</span><span class="sxs-lookup"><span data-stu-id="9229c-136">10/11/2019</span></span> |<span data-ttu-id="9229c-137">Crédit</span><span class="sxs-lookup"><span data-stu-id="9229c-137">Credit</span></span> |<span data-ttu-id="9229c-138">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="9229c-138">The Phone Company</span></span> |<span data-ttu-id="9229c-139">99,95</span><span class="sxs-lookup"><span data-stu-id="9229c-139">99.95</span></span> | |
    |<span data-ttu-id="9229c-140">13/10/2019</span><span class="sxs-lookup"><span data-stu-id="9229c-140">10/13/2019</span></span> |<span data-ttu-id="9229c-141">Crédit</span><span class="sxs-lookup"><span data-stu-id="9229c-141">Credit</span></span> |<span data-ttu-id="9229c-142">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="9229c-142">Coho Vineyard</span></span> |<span data-ttu-id="9229c-143">154,43</span><span class="sxs-lookup"><span data-stu-id="9229c-143">154.43</span></span> | |
    |<span data-ttu-id="9229c-144">15/10/2019</span><span class="sxs-lookup"><span data-stu-id="9229c-144">10/15/2019</span></span> |<span data-ttu-id="9229c-145">Compte courant</span><span class="sxs-lookup"><span data-stu-id="9229c-145">Checking</span></span> |<span data-ttu-id="9229c-146">Versement externe</span><span class="sxs-lookup"><span data-stu-id="9229c-146">External Deposit</span></span> | |<span data-ttu-id="9229c-147">1000</span><span class="sxs-lookup"><span data-stu-id="9229c-147">1000</span></span> |
    |<span data-ttu-id="9229c-148">20/10/2019</span><span class="sxs-lookup"><span data-stu-id="9229c-148">10/20/2019</span></span> |<span data-ttu-id="9229c-149">Crédit</span><span class="sxs-lookup"><span data-stu-id="9229c-149">Credit</span></span> |<span data-ttu-id="9229c-150">Coho Vineyard − Remboursement</span><span class="sxs-lookup"><span data-stu-id="9229c-150">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="9229c-151">−35,45</span><span class="sxs-lookup"><span data-stu-id="9229c-151">-35.45</span></span> |
    |<span data-ttu-id="9229c-152">25/10/2019</span><span class="sxs-lookup"><span data-stu-id="9229c-152">10/25/2019</span></span> |<span data-ttu-id="9229c-153">Compte courant</span><span class="sxs-lookup"><span data-stu-id="9229c-153">Checking</span></span> |<span data-ttu-id="9229c-154">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="9229c-154">Best For You Organics Company</span></span> | <span data-ttu-id="9229c-155">−85,64</span><span class="sxs-lookup"><span data-stu-id="9229c-155">-85.64</span></span> | |
    |<span data-ttu-id="9229c-156">01/11/2019</span><span class="sxs-lookup"><span data-stu-id="9229c-156">11/01/2019</span></span> |<span data-ttu-id="9229c-157">Compte courant</span><span class="sxs-lookup"><span data-stu-id="9229c-157">Checking</span></span> |<span data-ttu-id="9229c-158">Versement externe</span><span class="sxs-lookup"><span data-stu-id="9229c-158">External Deposit</span></span> | |<span data-ttu-id="9229c-159">1000</span><span class="sxs-lookup"><span data-stu-id="9229c-159">1000</span></span> |

3. <span data-ttu-id="9229c-160">Ouvrez l’**éditeur de code** puis sélectionnez **Nouveau script**.</span><span class="sxs-lookup"><span data-stu-id="9229c-160">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="9229c-161">Nous allons réarranger la mise en forme.</span><span class="sxs-lookup"><span data-stu-id="9229c-161">Let's clean up the formatting.</span></span> <span data-ttu-id="9229c-162">Il s’agit d’un document financier, nous allons donc modifier la mise en forme des nombres dans les colonnes **Débit** et **Crédit** pour afficher les valeurs sous forme de montants en dollars.</span><span class="sxs-lookup"><span data-stu-id="9229c-162">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="9229c-163">Ajustons également la largeur des colonnes aux données.</span><span class="sxs-lookup"><span data-stu-id="9229c-163">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="9229c-164">Remplacez le contenu du script par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="9229c-164">Replace the script contents with the following code:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. <span data-ttu-id="9229c-165">Nous allons maintenant lire une valeur depuis l’une des colonnes de montants.</span><span class="sxs-lookup"><span data-stu-id="9229c-165">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="9229c-166">Ajoutez le code suivant à la fin du script :</span><span class="sxs-lookup"><span data-stu-id="9229c-166">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    <span data-ttu-id="9229c-167">Remarquez les appels de `load` et de `sync`.</span><span class="sxs-lookup"><span data-stu-id="9229c-167">Note the calls to `load` and `sync`.</span></span> <span data-ttu-id="9229c-168">Pour plus de détails sur ces méthodes, voir [Principes de base des scripts Office dans Excel pour le web](../develop/scripting-fundamentals.md#sync-and-load).</span><span class="sxs-lookup"><span data-stu-id="9229c-168">You can learn the details of those methods in [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md#sync-and-load).</span></span> <span data-ttu-id="9229c-169">Pour l’instant, sachez seulement que vous devez demander la lecture des données puis synchroniser votre script avec le classeur pour lire les données.</span><span class="sxs-lookup"><span data-stu-id="9229c-169">For now, know that you must request data to be read and then sync your script with the workbook to read that data.</span></span>

6. <span data-ttu-id="9229c-170">Exécutez le script.</span><span class="sxs-lookup"><span data-stu-id="9229c-170">Run the script.</span></span>
7. <span data-ttu-id="9229c-171">Ouvrez la console.</span><span class="sxs-lookup"><span data-stu-id="9229c-171">Open the console.</span></span> <span data-ttu-id="9229c-172">Accédez au menu **Ellipses**, puis appuyez sur **Journaux...**.</span><span class="sxs-lookup"><span data-stu-id="9229c-172">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="9229c-173">Dans la console, `[Array[1]]` doit s’afficher.</span><span class="sxs-lookup"><span data-stu-id="9229c-173">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="9229c-174">Ce n’est pas un nombre, car les plages sont des tableaux de données à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="9229c-174">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="9229c-175">Cette plage à deux dimensions est directement journalisée dans la console.</span><span class="sxs-lookup"><span data-stu-id="9229c-175">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="9229c-176">Heureusement, l’éditeur de code vous permet de voir le contenu du tableau.</span><span class="sxs-lookup"><span data-stu-id="9229c-176">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="9229c-177">Lorsqu’un tableau à deux dimensions est journalisé sur la console, il regroupe les valeurs de colonne sous chaque ligne.</span><span class="sxs-lookup"><span data-stu-id="9229c-177">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="9229c-178">Développez le journal du tableau en appuyant sur le triangle bleu.</span><span class="sxs-lookup"><span data-stu-id="9229c-178">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="9229c-179">Développez le deuxième niveau du tableau en appuyant sur le triangle bleu nouvellement affiché.</span><span class="sxs-lookup"><span data-stu-id="9229c-179">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="9229c-180">Voici ce que vous devez voir :</span><span class="sxs-lookup"><span data-stu-id="9229c-180">You should now see this:</span></span>

    ![Journal de la console affichant la sortie « −20,05 », imbriquée sous deux tableaux.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="9229c-182">Modifier la valeur d’une cellule.</span><span class="sxs-lookup"><span data-stu-id="9229c-182">Modify the value of a cell</span></span>

<span data-ttu-id="9229c-183">Maintenant que nous avons vu comment lire des données, nous allons les utiliser pour modifier le classeur.</span><span class="sxs-lookup"><span data-stu-id="9229c-183">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="9229c-184">Nous allons rendre la valeur de la cellule **D2** positive avec la fonction `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="9229c-184">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="9229c-185">L’objet [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contient de nombreuses fonctions auxquelles vos scripts ont accès.</span><span class="sxs-lookup"><span data-stu-id="9229c-185">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="9229c-186">Pour plus d’informations sur `Math` et les autres objets intégrés, voir [Utilisation d’objets JavaScript intégrés dans les scripts Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="9229c-186">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="9229c-187">Ajoutez le code suivant à la fin du script :</span><span class="sxs-lookup"><span data-stu-id="9229c-187">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. <span data-ttu-id="9229c-188">La valeur de la cellule **D2** doit maintenant être positive.</span><span class="sxs-lookup"><span data-stu-id="9229c-188">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="9229c-189">Modifier les valeurs d’une colonne</span><span class="sxs-lookup"><span data-stu-id="9229c-189">Modify the values of a column</span></span>

<span data-ttu-id="9229c-190">Maintenant que nous avons vu comment lire et écrire dans une seule cellule, configurons le script de façon à ce qu’il travaille sur l’ensemble des cellules des colonnes **Débit** et **Crédit**.</span><span class="sxs-lookup"><span data-stu-id="9229c-190">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="9229c-191">Supprimez le code qui affecte une seule cellule (le code de valeur absolue précédent), de sorte que votre script se présente désormais comme suit :</span><span class="sxs-lookup"><span data-stu-id="9229c-191">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. <span data-ttu-id="9229c-192">Ajoutez une boucle pour produire une itération dans des lignes des deux dernières colonnes.</span><span class="sxs-lookup"><span data-stu-id="9229c-192">Add a loop that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="9229c-193">Le script remplace la valeur de chaque cellule en la valeur absolue de cette valeur.</span><span class="sxs-lookup"><span data-stu-id="9229c-193">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="9229c-194">Notez que l’indexation du tableau qui définit les emplacements des cellules est basée sur zéro.</span><span class="sxs-lookup"><span data-stu-id="9229c-194">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="9229c-195">Par conséquent, la cellule **A1** est `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="9229c-195">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    <span data-ttu-id="9229c-196">Cette partie du script effectue plusieurs tâches importantes.</span><span class="sxs-lookup"><span data-stu-id="9229c-196">This portion of the script does several important tasks.</span></span> <span data-ttu-id="9229c-197">Premièrement, elle charge les valeurs et le nombre de lignes de la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="9229c-197">First, it loads the values and row count of the used range.</span></span> <span data-ttu-id="9229c-198">Nous pouvons ainsi examiner les valeurs et déterminer quand arrêter.</span><span class="sxs-lookup"><span data-stu-id="9229c-198">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="9229c-199">Deuxièmement, elle produit une itération dans la plage utilisée, en vérifiant chaque cellule des colonnes **Débit** et **Crédit**.</span><span class="sxs-lookup"><span data-stu-id="9229c-199">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="9229c-200">Enfin, si la valeur dans la cellule n’est pas 0, elle est remplacée par sa valeur absolue.</span><span class="sxs-lookup"><span data-stu-id="9229c-200">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="9229c-201">Nous évitons les zéros pour pouvoir laisser les cellules vides telles qu’elles sont.</span><span class="sxs-lookup"><span data-stu-id="9229c-201">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="9229c-202">Exécutez le script.</span><span class="sxs-lookup"><span data-stu-id="9229c-202">Run the script.</span></span>

    <span data-ttu-id="9229c-203">Voici ce à quoi doit maintenant ressembler le relevé bancaire :</span><span class="sxs-lookup"><span data-stu-id="9229c-203">Your banking statement should now look like this:</span></span>

    ![Le relevé bancaire sous la forme d’un tableau mis en forme avec uniquement des valeurs positives.](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="9229c-205">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="9229c-205">Next steps</span></span>

<span data-ttu-id="9229c-206">Ouvrez l’éditeur de code et testez quelques-uns de nos [Exemples de scripts pour Scripts Office dans Excel pour le web](../resources/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="9229c-206">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="9229c-207">Vous pouvez également consulter [Principes de base des scripts Office dans Excel pour le web](../develop/scripting-fundamentals.md) pour en savoir plus sur la création de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="9229c-207">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
