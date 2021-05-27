---
title: Enregistrer, modifier, créer des scripts Office dans Excel pour le web
description: Didacticiel sur les notions de base des scripts Office, comprenant l’enregistrement de scripts avec l’enregistreur d’actions et l’écriture de données dans un classeur.
ms.date: 05/23/2021
localization_priority: Priority
ms.openlocfilehash: 19cd7bf6c3120d674553d37a36f45d36f46ee852
ms.sourcegitcommit: 0343e4a9843f7ab6ec99d6ddf955050271b061c7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/26/2021
ms.locfileid: "52655904"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="1a4cf-103">Enregistrer, modifier, créer des scripts Office dans Excel pour le web</span><span class="sxs-lookup"><span data-stu-id="1a4cf-103">Record, edit, and create Office Scripts in Excel on the web</span></span>

<span data-ttu-id="1a4cf-104">Ce didacticiel vous présente les notions de base de l’enregistrement, de la modification et de la rédaction d’un script Office pour Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-104">This tutorial teaches you the basics of recording, editing, and writing an Office Script for Excel on the web.</span></span> <span data-ttu-id="1a4cf-105">Vous allez enregistrer un script mettant en forme une feuille de calcul d’enregistrement des ventes.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-105">You'll record a script that applies some formatting to a sales record worksheet.</span></span> <span data-ttu-id="1a4cf-106">Vous allez ensuite modifier le script enregistré pour appliquer une mise en forme supplémentaire, créer un tableau, puis trier ce tableau.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-106">You'll then edit the recorded script to apply more formatting, create a table, and sort that table.</span></span> <span data-ttu-id="1a4cf-107">Ce modèle de type « enregistrement suivi d’une modification » constitue un outil important pour vous permettre de savoir à quoi ressemblent vos actions Excel sous forme de code.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-107">This record-then-edit pattern is an important tool to see what your Excel actions look like as code.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1a4cf-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="1a4cf-108">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="1a4cf-109">Ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-109">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="1a4cf-110">Si vous découvrez JavaScript, nous vous conseillons de commencer par consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="1a4cf-110">If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="1a4cf-111">Si vous souhaitez en savoir plus sur l’environnement de script, veuillez consulter la rubrique [Environnement de l’éditeur de code Scripts Office](../overview/code-editor-environment.md).</span><span class="sxs-lookup"><span data-stu-id="1a4cf-111">Visit [Office Scripts Code Editor environment](../overview/code-editor-environment.md) to learn more about the script environment.</span></span>

## <a name="add-data-and-record-a-basic-script"></a><span data-ttu-id="1a4cf-112">Ajouter des données et enregistrer un script simple</span><span class="sxs-lookup"><span data-stu-id="1a4cf-112">Add data and record a basic script</span></span>

<span data-ttu-id="1a4cf-113">Tout d’abord, il nous faut des données et un petit script de base.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-113">First, we'll need some data and a small starting script.</span></span>

1. <span data-ttu-id="1a4cf-114">Créez un nouveau classeur dans Excel pour le Web.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-114">Create a new workbook in Excel for the Web.</span></span>
2. <span data-ttu-id="1a4cf-115">Copiez les données de ventes de fruits suivantes et collez-les dans la feuille de calcul en commençant à la cellule **A1**.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-115">Copy the following fruit sales data and paste it into the worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="1a4cf-116">Fruits</span><span class="sxs-lookup"><span data-stu-id="1a4cf-116">Fruit</span></span> |<span data-ttu-id="1a4cf-117">2018</span><span class="sxs-lookup"><span data-stu-id="1a4cf-117">2018</span></span> |<span data-ttu-id="1a4cf-118">2019</span><span class="sxs-lookup"><span data-stu-id="1a4cf-118">2019</span></span> |
    |:---|:---|:---|
    |<span data-ttu-id="1a4cf-119">Oranges</span><span class="sxs-lookup"><span data-stu-id="1a4cf-119">Oranges</span></span> |<span data-ttu-id="1a4cf-120">1000</span><span class="sxs-lookup"><span data-stu-id="1a4cf-120">1000</span></span> |<span data-ttu-id="1a4cf-121">1200</span><span class="sxs-lookup"><span data-stu-id="1a4cf-121">1200</span></span> |
    |<span data-ttu-id="1a4cf-122">Citrons</span><span class="sxs-lookup"><span data-stu-id="1a4cf-122">Lemons</span></span> |<span data-ttu-id="1a4cf-123">800</span><span class="sxs-lookup"><span data-stu-id="1a4cf-123">800</span></span> |<span data-ttu-id="1a4cf-124">900</span><span class="sxs-lookup"><span data-stu-id="1a4cf-124">900</span></span> |
    |<span data-ttu-id="1a4cf-125">Citrons verts</span><span class="sxs-lookup"><span data-stu-id="1a4cf-125">Limes</span></span> |<span data-ttu-id="1a4cf-126">600</span><span class="sxs-lookup"><span data-stu-id="1a4cf-126">600</span></span> |<span data-ttu-id="1a4cf-127">500</span><span class="sxs-lookup"><span data-stu-id="1a4cf-127">500</span></span> |
    |<span data-ttu-id="1a4cf-128">Pamplemousses</span><span class="sxs-lookup"><span data-stu-id="1a4cf-128">Grapefruits</span></span> |<span data-ttu-id="1a4cf-129">900</span><span class="sxs-lookup"><span data-stu-id="1a4cf-129">900</span></span> |<span data-ttu-id="1a4cf-130">700</span><span class="sxs-lookup"><span data-stu-id="1a4cf-130">700</span></span> |

3. <span data-ttu-id="1a4cf-131">Ouvrez l’onglet **Automatiser**. Si vous ne voyez pas l’onglet **Automatiser**, vérifiez dans la section dépassement du ruban en appuyant sur la flèche déroulante vers le bas.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-131">Open the **Automate** tab. If you don't see the **Automate** tab, check the ribbon overflow by pressing the drop-down arrow.</span></span> <span data-ttu-id="1a4cf-132">Si vous ne voyez toujours pas l’onglet, suivez les conseils de l’article [Résolution des problèmes de scripts Office](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span><span class="sxs-lookup"><span data-stu-id="1a4cf-132">If it's still not there, follow the advice in the article [Troubleshoot Office Scripts](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable).</span></span>
4. <span data-ttu-id="1a4cf-133">Appuyez sur le bouton **Actions d’enregistrement**.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-133">Press the **Record Actions** button.</span></span>
5. <span data-ttu-id="1a4cf-134">Sélectionnez les cellules **A2:C2** (la ligne « Oranges ») et choisissez orange comme couleur de remplissage.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-134">Select cells **A2:C2** (the "Oranges" row) and set the fill color to orange.</span></span>
6. <span data-ttu-id="1a4cf-135">Appuyez sur le bouton **Arrêter** pour arrêter l’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-135">Stop the recording by pressing the **Stop** button.</span></span>

    <span data-ttu-id="1a4cf-136">Voici ce à quoi votre feuille de calcul doit ressembler (les couleurs peuvent être différentes) :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-136">Your worksheet should look like this (don't worry if the color is different):</span></span>

    :::image type="content" source="../images/tutorial-1.png" alt-text="Feuille de calcul affichant une ligne de données de ventes de fruits avec les « Oranges » mises en évidence par la couleur orange.":::

## <a name="edit-an-existing-script"></a><span data-ttu-id="1a4cf-138">Modifier un script existant</span><span class="sxs-lookup"><span data-stu-id="1a4cf-138">Edit an existing script</span></span>

<span data-ttu-id="1a4cf-139">Le script précédent a coloré la ligne « Oranges » en orange.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-139">The previous script colored the "Oranges" row to be orange.</span></span> <span data-ttu-id="1a4cf-140">Nous allons ajouter une ligne jaune pour « Citrons ».</span><span class="sxs-lookup"><span data-stu-id="1a4cf-140">Let's add a yellow row for the "Lemons".</span></span>

1. <span data-ttu-id="1a4cf-141">Depuis le volet **Détails** à présent ouvert, appuyez sur le bouton **Modifier**.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-141">From the now-open **Details** pane, press the **Edit** button.</span></span>
2. <span data-ttu-id="1a4cf-142">Un code similaire à celui-ci doit apparaître :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-142">You should see something similar to this code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    <span data-ttu-id="1a4cf-143">Ce code extrait la feuille de calcul actuelle du classeur.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-143">This code gets the current worksheet from the workbook.</span></span> <span data-ttu-id="1a4cf-144">Il définit ensuite la couleur de remplissage de la plage **A2:C2**.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-144">Then, it sets the fill color of the range **A2:C2**.</span></span>

    <span data-ttu-id="1a4cf-145">Les plages jouent un rôle fondamental dans les scripts Office d’Excel pour le web.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-145">Ranges are a fundamental part of Office Scripts in Excel on the web.</span></span> <span data-ttu-id="1a4cf-146">Une plage est un bloc de cellules contiguës de forme rectangulaire qui contient des valeurs, des formules ou des formats.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-146">A range is a contiguous, rectangular block of cells that contains values, formula, and formatting.</span></span> <span data-ttu-id="1a4cf-147">Les plages constituent la structure de base faite de cellules par laquelle vous effectuerez des tâches de script.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-147">They are the basic structure of cells through which you'll perform most of your scripting tasks.</span></span>

3. <span data-ttu-id="1a4cf-148">Ajoutez la ligne suivante à la fin du script (entre l’emplacement où le `color` se trouve et le `}` de clôture) :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-148">Add the following line to the end of the script (between where the `color` is set and the closing `}`):</span></span>

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. <span data-ttu-id="1a4cf-149">Testez le script en appuyant sur **Exécuter**.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-149">Test the script by pressing **Run**.</span></span> <span data-ttu-id="1a4cf-150">Voici ce à quoi votre feuille de calcul doit maintenant ressembler :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-150">Your workbook should now look like this:</span></span>

    :::image type="content" source="../images/tutorial-2.png" alt-text="Feuille de calcul affichant la ligne des données de ventes de fruits avec la ligne « Oranges » mise en évidence par la couleur orange et la ligne « Citrons » par la couleur jaune.":::

## <a name="create-a-table"></a><span data-ttu-id="1a4cf-152">Créer un tableau</span><span class="sxs-lookup"><span data-stu-id="1a4cf-152">Create a table</span></span>

<span data-ttu-id="1a4cf-153">Nous allons convertir les données de ventes de fruits en tableau.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-153">Let's convert this fruit sales data into a table.</span></span> <span data-ttu-id="1a4cf-154">Nous allons utiliser notre script pour l’ensemble du processus.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-154">We'll use our script for the entire process.</span></span>

1. <span data-ttu-id="1a4cf-155">Ajoutez la ligne suivante à la fin du script (avant le `}` de clôture) :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-155">Add the following line to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. <span data-ttu-id="1a4cf-156">Cet appel renvoie un objet `Table`.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-156">That call returns a `Table` object.</span></span> <span data-ttu-id="1a4cf-157">Nous allons utiliser ce tableau pour trier les données.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-157">Let's use that table to sort the data.</span></span> <span data-ttu-id="1a4cf-158">Nous allons trier les données en ordre croissant en fonction des valeurs de la colonne « Fruits ».</span><span class="sxs-lookup"><span data-stu-id="1a4cf-158">We'll sort the data in ascending order based on the values in the "Fruit" column.</span></span> <span data-ttu-id="1a4cf-159">Ajoutez la ligne suivante après la création du tableau :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-159">Add the following line after the table creation:</span></span>

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    <span data-ttu-id="1a4cf-160">Voici ce à quoi doit ressembler votre script :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-160">Your script should look like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    <span data-ttu-id="1a4cf-161">Les tableaux comportent un objet `TableSort`, accessible via la méthode `Table.getSort`.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-161">Tables have a `TableSort` object, accessed through the `Table.getSort` method.</span></span> <span data-ttu-id="1a4cf-162">Vous pouvez appliquer des critères de tri à cet objet.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-162">You can apply sorting criteria to that object.</span></span> <span data-ttu-id="1a4cf-163">La méthode `apply` prend un tableau d’objets `SortField`.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-163">The `apply` method takes in an array of `SortField` objects.</span></span> <span data-ttu-id="1a4cf-164">Dans notre cas, ne disposant que d’un seul critère de tri, nous utiliserons un seul `SortField`.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-164">In this case, we only have one sorting criteria, so we only use one `SortField`.</span></span> <span data-ttu-id="1a4cf-165">`key: 0` définit la colonne ayant les valeurs de définition de tri sur « 0 » (la première colonne du tableau, **A** dans notre cas).</span><span class="sxs-lookup"><span data-stu-id="1a4cf-165">`key: 0` sets the column with the sort-defining values to "0" (which is the first column on the table, **A** in this case).</span></span> <span data-ttu-id="1a4cf-166">`ascending: true` trie les données dans un ordre croissant (et non dans un ordre décroissant).</span><span class="sxs-lookup"><span data-stu-id="1a4cf-166">`ascending: true` sorts the data in ascending order (instead of descending order).</span></span>

3. <span data-ttu-id="1a4cf-p111">Exécutez le script. Un tableau come ceci devrait s’afficher :</span><span class="sxs-lookup"><span data-stu-id="1a4cf-p111">Run the script. You should see a table like this:</span></span>

    :::image type="content" source="../images/tutorial-3.png" alt-text="Feuille de calcul affichant la table de ventes des fruits triées.":::

    > [!NOTE]
    > <span data-ttu-id="1a4cf-170">Si vous réexécutez le script, un message d’erreur s’affiche.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-170">If you re-run the script, you'll get an error.</span></span> <span data-ttu-id="1a4cf-171">En effet, vous ne pouvez pas créer un tableau au-dessus d’un autre.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-171">This is because you cannot create a table on top of another table.</span></span> <span data-ttu-id="1a4cf-172">Toutefois, vous pouvez exécuter le script sur une autre feuille de calcul ou un autre classeur.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-172">However, you can run the script on a different worksheet or workbook.</span></span>

### <a name="re-run-the-script"></a><span data-ttu-id="1a4cf-173">Réexécutez le script.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-173">Re-run the script</span></span>

1. <span data-ttu-id="1a4cf-174">Créer une nouvelle feuille de calcul dans le classeur actif.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-174">Create a new worksheet in the current workbook.</span></span>
2. <span data-ttu-id="1a4cf-175">Copiez les données des fruits du début de ce didacticiel et collez-les dans la nouvelle feuille de calcul, en commençant à la cellule **A1**.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-175">Copy the fruit data from the beginning of the tutorial and paste it into the new worksheet, starting at cell **A1**.</span></span>
3. <span data-ttu-id="1a4cf-176">Exécutez le script.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-176">Run the script.</span></span>

## <a name="next-steps"></a><span data-ttu-id="1a4cf-177">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="1a4cf-177">Next steps</span></span>

<span data-ttu-id="1a4cf-178">Complétez le didacticiel [Lire les données d’un classeur avec les scripts Office d’Excel pour le web](excel-read-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="1a4cf-178">Complete the [Read workbook data with Office Scripts in Excel on the web](excel-read-tutorial.md) tutorial.</span></span> <span data-ttu-id="1a4cf-179">Il vous apprend comment lire des données à partir d’un classeur à l’aide d’un script Office.</span><span class="sxs-lookup"><span data-stu-id="1a4cf-179">It teaches you how to read data from a workbook with an Office Script.</span></span>
