---
title: Lire les données d’un classeur avec les scripts Office d’Excel pour le web
description: Didacticiel des scripts Office sur la lecture de données à partir de classeurs et l’évaluation de ces données dans le script.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: aa05533f0d7cd3b53e4eb616ae3216d672d026f9
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074689"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="01794-103">Lire les données d’un classeur avec les scripts Office d’Excel pour le web</span><span class="sxs-lookup"><span data-stu-id="01794-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="01794-104">Ce didacticiel vous apprend comment lire des données à partir d’un classeur à l’aide d’un script Office pour Excel pour le web.</span><span class="sxs-lookup"><span data-stu-id="01794-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="01794-105">Vous allez écrire un nouveau script qui met en forme un relevé bancaire et normalise les données incluses.</span><span class="sxs-lookup"><span data-stu-id="01794-105">You'll be writing a new script that formats a bank statement and normalizes the data in that statement.</span></span> <span data-ttu-id="01794-106">Lors de ce nettoyage de données, votre script lira les valeurs des cellules de transaction, appliquera une formule simple à chaque valeur, puis écrira la réponse résultante dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="01794-106">As part of that data clean-up, your script will read values from the transaction cells, apply a simple formula to each value, and write the resulting answer to the workbook.</span></span> <span data-ttu-id="01794-107">La lecture de données du classeur vous permet d’automatiser certains processus décisionnels dans le script.</span><span class="sxs-lookup"><span data-stu-id="01794-107">Reading data from the workbook lets you automate some of your decision making processes in the script.</span></span>

> [!TIP]
> <span data-ttu-id="01794-108">Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="01794-108">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span> <span data-ttu-id="01794-109">[Les scripts Office utilisent TypeScript](../overview/code-editor-environment.md), et ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="01794-109">[Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="01794-110">Si vous découvrez JavaScript, nous vous conseillons de commencer par consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="01794-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="01794-111">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="01794-111">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a><span data-ttu-id="01794-112">Lire une cellule</span><span class="sxs-lookup"><span data-stu-id="01794-112">Read a cell</span></span>

<span data-ttu-id="01794-113">Les scripts créés avec l’enregistreur d’actions peuvent uniquement écrire des informations dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="01794-113">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="01794-114">L’éditeur de code vous permet de modifier et de créer des scripts qui peuvent également lire les données d’un classeur.</span><span class="sxs-lookup"><span data-stu-id="01794-114">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="01794-115">Nous allons créer un script qui lit les données et agit en fonction de ce qui a été lu.</span><span class="sxs-lookup"><span data-stu-id="01794-115">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="01794-116">Nous allons utiliser un exemple de relevé bancaire.</span><span class="sxs-lookup"><span data-stu-id="01794-116">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="01794-117">Il s’agit d’un relevé combiné de compte courant et de crédit.</span><span class="sxs-lookup"><span data-stu-id="01794-117">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="01794-118">Malheureusement, les changements de soldes sont rapportés différemment.</span><span class="sxs-lookup"><span data-stu-id="01794-118">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="01794-119">Le relevé de compte courant donne les revenus comme crédit positif et les dépenses comme débit négatif.</span><span class="sxs-lookup"><span data-stu-id="01794-119">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="01794-120">Le relevé de crédit fait l’inverse.</span><span class="sxs-lookup"><span data-stu-id="01794-120">The credit statement does the opposite.</span></span>

<span data-ttu-id="01794-121">Dans le reste du didacticiel, nous allons normaliser ces données à l’aide d’un script.</span><span class="sxs-lookup"><span data-stu-id="01794-121">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="01794-122">Pour commencer, voyons comment lire des données à partir du classeur.</span><span class="sxs-lookup"><span data-stu-id="01794-122">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="01794-123">Créez une nouvelle feuille de calcul dans le classeur courant, vous l’utiliserez pour le reste du didacticiel.</span><span class="sxs-lookup"><span data-stu-id="01794-123">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="01794-124">Copiez les données suivantes et collez-les dans la feuille de calcul en commençant à la cellule **A1**.</span><span class="sxs-lookup"><span data-stu-id="01794-124">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="01794-125">Date</span><span class="sxs-lookup"><span data-stu-id="01794-125">Date</span></span> |<span data-ttu-id="01794-126">Compte</span><span class="sxs-lookup"><span data-stu-id="01794-126">Account</span></span> |<span data-ttu-id="01794-127">Description</span><span class="sxs-lookup"><span data-stu-id="01794-127">Description</span></span> |<span data-ttu-id="01794-128">Débit</span><span class="sxs-lookup"><span data-stu-id="01794-128">Debit</span></span> |<span data-ttu-id="01794-129">Crédit</span><span class="sxs-lookup"><span data-stu-id="01794-129">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="01794-130">10/10/2019</span><span class="sxs-lookup"><span data-stu-id="01794-130">10/10/2019</span></span> |<span data-ttu-id="01794-131">Compte courant</span><span class="sxs-lookup"><span data-stu-id="01794-131">Checking</span></span> |<span data-ttu-id="01794-132">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="01794-132">Coho Vineyard</span></span> |<span data-ttu-id="01794-133">−20,05</span><span class="sxs-lookup"><span data-stu-id="01794-133">-20.05</span></span> | |
    |<span data-ttu-id="01794-134">11/10/2019</span><span class="sxs-lookup"><span data-stu-id="01794-134">10/11/2019</span></span> |<span data-ttu-id="01794-135">Crédit</span><span class="sxs-lookup"><span data-stu-id="01794-135">Credit</span></span> |<span data-ttu-id="01794-136">The Phone Company</span><span class="sxs-lookup"><span data-stu-id="01794-136">The Phone Company</span></span> |<span data-ttu-id="01794-137">99,95</span><span class="sxs-lookup"><span data-stu-id="01794-137">99.95</span></span> | |
    |<span data-ttu-id="01794-138">13/10/2019</span><span class="sxs-lookup"><span data-stu-id="01794-138">10/13/2019</span></span> |<span data-ttu-id="01794-139">Crédit</span><span class="sxs-lookup"><span data-stu-id="01794-139">Credit</span></span> |<span data-ttu-id="01794-140">Coho Vineyard</span><span class="sxs-lookup"><span data-stu-id="01794-140">Coho Vineyard</span></span> |<span data-ttu-id="01794-141">154,43</span><span class="sxs-lookup"><span data-stu-id="01794-141">154.43</span></span> | |
    |<span data-ttu-id="01794-142">15/10/2019</span><span class="sxs-lookup"><span data-stu-id="01794-142">10/15/2019</span></span> |<span data-ttu-id="01794-143">Compte courant</span><span class="sxs-lookup"><span data-stu-id="01794-143">Checking</span></span> |<span data-ttu-id="01794-144">Versement externe</span><span class="sxs-lookup"><span data-stu-id="01794-144">External Deposit</span></span> | |<span data-ttu-id="01794-145">1000</span><span class="sxs-lookup"><span data-stu-id="01794-145">1000</span></span> |
    |<span data-ttu-id="01794-146">20/10/2019</span><span class="sxs-lookup"><span data-stu-id="01794-146">10/20/2019</span></span> |<span data-ttu-id="01794-147">Crédit</span><span class="sxs-lookup"><span data-stu-id="01794-147">Credit</span></span> |<span data-ttu-id="01794-148">Coho Vineyard − Remboursement</span><span class="sxs-lookup"><span data-stu-id="01794-148">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="01794-149">−35,45</span><span class="sxs-lookup"><span data-stu-id="01794-149">-35.45</span></span> |
    |<span data-ttu-id="01794-150">25/10/2019</span><span class="sxs-lookup"><span data-stu-id="01794-150">10/25/2019</span></span> |<span data-ttu-id="01794-151">Compte courant</span><span class="sxs-lookup"><span data-stu-id="01794-151">Checking</span></span> |<span data-ttu-id="01794-152">Best For You Organics Company</span><span class="sxs-lookup"><span data-stu-id="01794-152">Best For You Organics Company</span></span> | <span data-ttu-id="01794-153">−85,64</span><span class="sxs-lookup"><span data-stu-id="01794-153">-85.64</span></span> | |
    |<span data-ttu-id="01794-154">01/11/2019</span><span class="sxs-lookup"><span data-stu-id="01794-154">11/01/2019</span></span> |<span data-ttu-id="01794-155">Compte courant</span><span class="sxs-lookup"><span data-stu-id="01794-155">Checking</span></span> |<span data-ttu-id="01794-156">Versement externe</span><span class="sxs-lookup"><span data-stu-id="01794-156">External Deposit</span></span> | |<span data-ttu-id="01794-157">1000</span><span class="sxs-lookup"><span data-stu-id="01794-157">1000</span></span> |

3. <span data-ttu-id="01794-158">Ouvrez **Tous les scripts** et sélectionner **Nouveau script**.</span><span class="sxs-lookup"><span data-stu-id="01794-158">Open **All Scripts** and select **New Script**.</span></span>
4. <span data-ttu-id="01794-159">Nous allons réarranger la mise en forme.</span><span class="sxs-lookup"><span data-stu-id="01794-159">Let's clean up the formatting.</span></span> <span data-ttu-id="01794-160">Il s’agit d’un document financier, nous allons donc modifier la mise en forme des nombres dans les colonnes **Débit** et **Crédit** pour afficher les valeurs sous forme de montants en dollars.</span><span class="sxs-lookup"><span data-stu-id="01794-160">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="01794-161">Ajustons également la largeur des colonnes aux données.</span><span class="sxs-lookup"><span data-stu-id="01794-161">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="01794-162">Remplacez le contenu du script par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="01794-162">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="01794-163">Nous allons maintenant lire une valeur depuis l’une des colonnes de montants.</span><span class="sxs-lookup"><span data-stu-id="01794-163">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="01794-164">Ajoutez le code suivant à la fin du script (avant le `}` de clôture) :</span><span class="sxs-lookup"><span data-stu-id="01794-164">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="01794-165">Exécutez le script.</span><span class="sxs-lookup"><span data-stu-id="01794-165">Run the script.</span></span>
7. <span data-ttu-id="01794-166">Dans la console, `[Array[1]]` doit s’afficher.</span><span class="sxs-lookup"><span data-stu-id="01794-166">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="01794-167">Ce n’est pas un nombre, car les plages sont des tableaux de données à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="01794-167">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="01794-168">Cette plage à deux dimensions est directement journalisée dans la console.</span><span class="sxs-lookup"><span data-stu-id="01794-168">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="01794-169">Heureusement, l’éditeur de code vous permet de voir le contenu du tableau.</span><span class="sxs-lookup"><span data-stu-id="01794-169">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="01794-170">Lorsqu’un tableau à deux dimensions est journalisé sur la console, il regroupe les valeurs de colonne sous chaque ligne.</span><span class="sxs-lookup"><span data-stu-id="01794-170">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="01794-171">Développez le journal du tableau en appuyant sur le triangle bleu.</span><span class="sxs-lookup"><span data-stu-id="01794-171">Expand the array log by pressing the blue triangle.</span></span>
9. <span data-ttu-id="01794-p110">Développez le deuxième niveau du tableau en appuyant sur le triangle bleu nouvellement révélé. Vous devriez maintenant voir ceci :</span><span class="sxs-lookup"><span data-stu-id="01794-p110">Expand the second level of the array by pressing the newly revealed blue triangle. You should now see this:</span></span>

    :::image type="content" source="../images/tutorial-4.png" alt-text="Journal de la console affichant la sortie « −20,05 », imbriquée dans deux tableaux.":::

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="01794-175">Modifier la valeur d’une cellule.</span><span class="sxs-lookup"><span data-stu-id="01794-175">Modify the value of a cell</span></span>

<span data-ttu-id="01794-176">Maintenant que nous avons vu comment lire des données, nous allons les utiliser pour modifier le classeur.</span><span class="sxs-lookup"><span data-stu-id="01794-176">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="01794-177">Nous allons rendre la valeur de la cellule **D2** positive avec la fonction `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="01794-177">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="01794-178">L’objet [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contient de nombreuses fonctions auxquelles vos scripts ont accès.</span><span class="sxs-lookup"><span data-stu-id="01794-178">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="01794-179">Pour plus d’informations sur `Math` et les autres objets intégrés, voir [Utilisation d’objets JavaScript intégrés dans les scripts Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="01794-179">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="01794-180">Nous utiliserons les méthodes `getValue` et `setValue` pour modifier la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="01794-180">We'll use `getValue` and `setValue` methods to change the value of the cell.</span></span> <span data-ttu-id="01794-181">Ces méthodes fonctionnent sur une seule cellule.</span><span class="sxs-lookup"><span data-stu-id="01794-181">These methods work on a single cell.</span></span> <span data-ttu-id="01794-182">Lorsque vous manipulez des plages de plusieurs cellules, vous pouvez utiliser `getValues` et `setValues`.</span><span class="sxs-lookup"><span data-stu-id="01794-182">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span> <span data-ttu-id="01794-183">Ajoutez le code suivant à la fin du script :</span><span class="sxs-lookup"><span data-stu-id="01794-183">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > <span data-ttu-id="01794-184">Nous [transformons](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) la valeur retournée de `range.getValue()` en `number` à l'aide du mot-clé `as`.</span><span class="sxs-lookup"><span data-stu-id="01794-184">We are [casting](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) the returned value of `range.getValue()` to a `number` by using the `as` keyword.</span></span> <span data-ttu-id="01794-185">Ceci est nécessaire, car une plage peut être des chaînes, des nombres ou des valeurs booléennes.</span><span class="sxs-lookup"><span data-stu-id="01794-185">This is necessary because a range could be strings, numbers, or booleans.</span></span> <span data-ttu-id="01794-186">Dans ce cas, nous avons explicitement besoin d’un nombre.</span><span class="sxs-lookup"><span data-stu-id="01794-186">In this instance, we explicitly need a number.</span></span>

2. <span data-ttu-id="01794-187">La valeur de la cellule **D2** doit maintenant être positive.</span><span class="sxs-lookup"><span data-stu-id="01794-187">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="01794-188">Modifier les valeurs d’une colonne</span><span class="sxs-lookup"><span data-stu-id="01794-188">Modify the values of a column</span></span>

<span data-ttu-id="01794-189">Maintenant que nous avons vu comment lire et écrire dans une seule cellule, configurons le script de façon à ce qu’il travaille sur l’ensemble des cellules des colonnes **Débit** et **Crédit**.</span><span class="sxs-lookup"><span data-stu-id="01794-189">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="01794-190">Supprimez le code qui affecte une seule cellule (le code de valeur absolue précédent), de sorte que votre script se présente désormais comme suit :</span><span class="sxs-lookup"><span data-stu-id="01794-190">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="01794-191">Ajoutez une boucle à la fin du script qui itère au sein des lignes des deux dernières colonnes.</span><span class="sxs-lookup"><span data-stu-id="01794-191">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="01794-192">Le script remplace la valeur de chaque cellule en la valeur absolue de cette valeur.</span><span class="sxs-lookup"><span data-stu-id="01794-192">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="01794-193">Notez que l’indexation du tableau qui définit les emplacements des cellules est basée sur zéro.</span><span class="sxs-lookup"><span data-stu-id="01794-193">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="01794-194">Par conséquent, la cellule **A1** est `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="01794-194">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="01794-195">Cette partie du script effectue plusieurs tâches importantes.</span><span class="sxs-lookup"><span data-stu-id="01794-195">This portion of the script does several important tasks.</span></span> <span data-ttu-id="01794-196">Premièrement, elle obtient les valeurs et le nombre de lignes de la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="01794-196">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="01794-197">Nous pouvons ainsi examiner les valeurs et déterminer quand arrêter.</span><span class="sxs-lookup"><span data-stu-id="01794-197">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="01794-198">Deuxièmement, elle produit une itération dans la plage utilisée, en vérifiant chaque cellule des colonnes **Débit** et **Crédit**.</span><span class="sxs-lookup"><span data-stu-id="01794-198">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="01794-199">Enfin, si la valeur dans la cellule n’est pas 0, elle est remplacée par sa valeur absolue.</span><span class="sxs-lookup"><span data-stu-id="01794-199">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="01794-200">Nous évitons les zéros pour pouvoir laisser les cellules vides telles qu’elles sont.</span><span class="sxs-lookup"><span data-stu-id="01794-200">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="01794-201">Exécutez le script.</span><span class="sxs-lookup"><span data-stu-id="01794-201">Run the script.</span></span>

    <span data-ttu-id="01794-202">Voici ce à quoi doit maintenant ressembler le relevé bancaire :</span><span class="sxs-lookup"><span data-stu-id="01794-202">Your banking statement should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-5.png" alt-text="Une feuille de calcul affichant le relevé bancaire sous la forme d’un tableau mis en forme avec uniquement des valeurs positives":::

## <a name="next-steps"></a><span data-ttu-id="01794-204">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="01794-204">Next steps</span></span>

<span data-ttu-id="01794-205">Ouvrez l’éditeur de code et testez quelques-uns de nos [Exemples de scripts pour Scripts Office dans Excel pour le web](../resources/samples/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="01794-205">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/samples/excel-samples.md).</span></span> <span data-ttu-id="01794-206">Vous pouvez également consulter [Principes de base des scripts Office dans Excel pour le web](../develop/scripting-fundamentals.md) pour en savoir plus sur la création de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="01794-206">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>

<span data-ttu-id="01794-207">La prochaine série de didacticiels sur les scripts Office met l’accent sur l’utilisation de scripts Office avec Power Automate.</span><span class="sxs-lookup"><span data-stu-id="01794-207">The next series of Office Scripts tutorials focus on using Office Scripts with Power Automate.</span></span> <span data-ttu-id="01794-208">Si vous souhaitez en savoir plus sur les avantages de la combinaison des deux plateformes, veuillez consulter [Exécuter des scripts Office avec Power Automate](../develop/power-automate-integration.md). Vous pouvez également essayer le didacticiel [Appeler des scripts à partir d’un flux manuel Power Automate](excel-power-automate-manual.md) pour créer un flux Power Automate utilisant un script Office.</span><span class="sxs-lookup"><span data-stu-id="01794-208">Learn more about the advantages combining the two platforms in [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) or try the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial to create a Power Automate flow that uses an Office Script.</span></span>
