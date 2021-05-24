---
title: Meilleures pratiques dans Office scripts
description: Comment éviter les problèmes courants et écrire des Office scripts fiables qui peuvent gérer des données ou des entrées inattendues.
ms.date: 05/10/2021
localization_priority: Normal
ms.openlocfilehash: 0697e6fd1fa8f437a4a585d938254deb5a05f20c
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546024"
---
# <a name="best-practices-in-office-scripts"></a><span data-ttu-id="2d5b7-103">Meilleures pratiques dans Office scripts</span><span class="sxs-lookup"><span data-stu-id="2d5b7-103">Best practices in Office Scripts</span></span>

<span data-ttu-id="2d5b7-104">Ces modèles et pratiques sont conçus pour aider vos scripts à s’exécuter correctement à chaque fois.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-104">These patterns and practices are designed to help your scripts run successfully every time.</span></span> <span data-ttu-id="2d5b7-105">Utilisez-les pour éviter les pièges courants lorsque vous commencez à automatiser Excel flux de travail.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-105">Use them to avoid common pitfalls as you start automating your Excel workflow.</span></span>

## <a name="verify-an-object-is-present"></a><span data-ttu-id="2d5b7-106">Vérifier la présence d’un objet</span><span class="sxs-lookup"><span data-stu-id="2d5b7-106">Verify an object is present</span></span>

<span data-ttu-id="2d5b7-107">Les scripts s’appuient souvent sur une feuille de calcul ou une table en cours de présence dans le workbook.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-107">Scripts often rely on a certain worksheet or table being present in the workbook.</span></span> <span data-ttu-id="2d5b7-108">Toutefois, ils peuvent être renommés ou supprimés entre les séquences de script.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-108">However, they might get renamed or removed between script runs.</span></span> <span data-ttu-id="2d5b7-109">En vérifiant si ces tableaux ou feuilles de calcul existent avant d’y appeler des méthodes, vous pouvez vous assurer que le script ne se termine pas brusquement.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-109">By checking if those tables or worksheets exist before calling methods on them, you can make sure the script doesn't end abruptly.</span></span>

<span data-ttu-id="2d5b7-110">L’exemple de code suivant vérifie si la feuille de calcul « Index » est présente dans le manuel.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-110">The following sample code checks if the "Index" worksheet is present in the workbook.</span></span> <span data-ttu-id="2d5b7-111">Si la feuille de calcul est présente, le script obtient une plage et continue.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-111">If the worksheet is present, the script gets a range and proceeds.</span></span> <span data-ttu-id="2d5b7-112">S’il n’est pas présent, le script enregistre un message d’erreur personnalisé.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-112">If it isn't present, the script logs a custom error message.</span></span>

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

<span data-ttu-id="2d5b7-113">L’opérateur TypeScript `?` vérifie si l’objet existe avant d’appeler une méthode.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-113">The TypeScript `?` operator checks if the object exists before calling a method.</span></span> <span data-ttu-id="2d5b7-114">Cela peut simplifier votre code si vous n’avez rien de spécial à faire lorsque l’objet n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-114">This can make your code more streamlined if you don't need to do anything special when the object doesn't exist.</span></span>

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a><span data-ttu-id="2d5b7-115">Valider d’abord les données et l’état du workbook</span><span class="sxs-lookup"><span data-stu-id="2d5b7-115">Validate data and workbook state first</span></span>

<span data-ttu-id="2d5b7-116">Assurez-vous que toutes vos feuilles de calcul, tableaux, formes et autres objets sont présents avant de travailler sur les données.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-116">Make sure all your worksheets, tables, shapes, and other objects are present before working on the data.</span></span> <span data-ttu-id="2d5b7-117">À l’aide du modèle précédent, vérifiez si tout se trouve dans le workbook et correspond à vos attentes.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-117">Using the previous pattern, check to see if everything is in the workbook and matches your expectations.</span></span> <span data-ttu-id="2d5b7-118">Le fait de le faire avant l’écriture de données garantit que votre script ne laisse pas le workbook dans un état partiel.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-118">Doing this before any data is written ensures your script doesn't leave the workbook in a partial state.</span></span>

<span data-ttu-id="2d5b7-119">Le script suivant requiert la présence de deux tables nommées « Table1 » et « Table2 ».</span><span class="sxs-lookup"><span data-stu-id="2d5b7-119">The following script requires two tables named "Table1" and "Table2" to be present.</span></span> <span data-ttu-id="2d5b7-120">Le script vérifie d’abord si les tables sont présentes, puis se termine par l’instruction et un message approprié `return` si ce n’est pas le cas.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-120">The script first checks if the tables are present and then ends with the `return` statement and an appropriate message if they're not.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue....
}
```

<span data-ttu-id="2d5b7-121">Si la vérification se produit dans une fonction distincte, vous devez quand même mettre fin au script en émettant `return` l’instruction à partir de la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-121">If the verification is happening in a separate function, you still must end the script by issuing the `return` statement from the `main` function.</span></span> <span data-ttu-id="2d5b7-122">Le retour à partir de la sous-section ne termine pas le script.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-122">Returning from the subfunction doesn't end the script.</span></span>

<span data-ttu-id="2d5b7-123">Le script suivant a le même comportement que le précédent.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-123">The following script has the same behavior as the previous one.</span></span> <span data-ttu-id="2d5b7-124">La différence est que la `main` fonction appelle la fonction pour tout `inputPresent` vérifier.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-124">The difference is that the `main` function calls the `inputPresent` function to verify everything.</span></span> <span data-ttu-id="2d5b7-125">`inputPresent` renvoie un booléen ( `true` ou ) pour indiquer si toutes les `false` entrées requises sont présentes.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-125">`inputPresent` returns a boolean (`true` or `false`) to indicate whether all required inputs are present.</span></span> <span data-ttu-id="2d5b7-126">La `main` fonction utilise ce type booléen pour décider de poursuivre ou de mettre fin au script.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-126">The `main` function uses that boolean to decide on continuing or ending the script.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue....
}

function inputPresent( workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a><span data-ttu-id="2d5b7-127">Quand utiliser une `throw` instruction</span><span class="sxs-lookup"><span data-stu-id="2d5b7-127">When to use a `throw` statement</span></span>

<span data-ttu-id="2d5b7-128">Une [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) instruction indique qu’une erreur inattendue s’est produite.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-128">A [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) statement indicates an unexpected error has occurred.</span></span> <span data-ttu-id="2d5b7-129">Il termine immédiatement le code.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-129">It ends the code immediately.</span></span> <span data-ttu-id="2d5b7-130">En grande partie, vous n’avez pas besoin de `throw` le faire à partir de votre script.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-130">For the most part, you don't need to `throw` from your script.</span></span> <span data-ttu-id="2d5b7-131">En règle générale, le script informe automatiquement l’utilisateur que le script n’a pas réussi à s’exécuter en raison d’un problème.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-131">Usually, the script automatically informs the user that the script failed to run due to an issue.</span></span> <span data-ttu-id="2d5b7-132">Dans la plupart des cas, il suffit de terminer le script avec un message d’erreur et `return` une instruction de la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-132">In most cases, it's sufficient to end the script with an error message and a `return` statement from the `main` function.</span></span>

<span data-ttu-id="2d5b7-133">Toutefois, si votre script s’exécute dans le cadre d’un flux Power Automate, vous pouvez arrêter le flux de continuer.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-133">However, if your script is running as part of a Power Automate flow, you may want to stop the flow from continuing.</span></span> <span data-ttu-id="2d5b7-134">Une `throw` instruction arrête le script et indique au flux de s’arrêter également.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-134">A `throw` statement stops the script and tells the flow to stop as well.</span></span>

<span data-ttu-id="2d5b7-135">Le script suivant montre comment utiliser `throw` l’instruction dans notre exemple de vérification de table.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-135">The following script shows how to use the `throw` statement in our table checking example.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a><span data-ttu-id="2d5b7-136">Quand utiliser une `try...catch` instruction</span><span class="sxs-lookup"><span data-stu-id="2d5b7-136">When to use a `try...catch` statement</span></span>

<span data-ttu-id="2d5b7-137">L’instruction permet de détecter si un appel [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) d’API échoue et de continuer à l’exécution du script.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-137">The [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statement is a way to detect if an API call fails and continue running the script.</span></span>

<span data-ttu-id="2d5b7-138">Prenons l’extrait de code suivant qui effectue une mise à jour de données importante sur une plage.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-138">Consider the following snippet that performs a large data update on a range.</span></span>

```TypeScript
range.setValues(someLargeValues);
```

<span data-ttu-id="2d5b7-139">Si `someLargeValues` la taille est supérieure à Excel sur le Web, l’appel `setValues()` échoue.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-139">If `someLargeValues` is larger than Excel for the web can handle, the `setValues()` call fails.</span></span> <span data-ttu-id="2d5b7-140">Le script échoue également avec une erreur [d’runtime.](../testing/troubleshooting.md#runtime-errors)</span><span class="sxs-lookup"><span data-stu-id="2d5b7-140">The script then also fails with a [runtime error](../testing/troubleshooting.md#runtime-errors).</span></span> <span data-ttu-id="2d5b7-141">L’instruction permet à votre script de reconnaître cette condition, sans terminer immédiatement le `try...catch` script et afficher l’erreur par défaut.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-141">The `try...catch` statement lets your script recognize this condition, without immediately ending the script and showing the default error.</span></span>

<span data-ttu-id="2d5b7-142">Une approche pour offrir à l’utilisateur du script une meilleure expérience consiste à lui présenter un message d’erreur personnalisé.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-142">One approach for giving the script user a better experience is to present them a custom error message.</span></span> <span data-ttu-id="2d5b7-143">L’extrait de code suivant montre une instruction consignant plus d’informations sur les `try...catch` erreurs pour mieux aider le lecteur.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-143">The following snippet shows a `try...catch` statement logging more error information to better help the reader.</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

<span data-ttu-id="2d5b7-144">Une autre approche de traitement des erreurs consiste à avoir un comportement de retour qui gère le cas d’erreur.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-144">Another approach to dealing with errors is to have fallback behavior that handles the error case.</span></span> <span data-ttu-id="2d5b7-145">L’extrait de code suivant utilise le bloc pour essayer une autre méthode qui décompose la mise à jour en plus petites parties `catch` et évite l’erreur.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-145">The following snippet uses the `catch` block to try an alternate method break up the update into smaller pieces and avoid the error.</span></span>

> [!TIP]
> <span data-ttu-id="2d5b7-146">Pour obtenir un exemple complet sur la mise à jour d’une grande plage, voir [Écrire un jeu de données de grande taille.](../resources/samples/write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="2d5b7-146">For a full example on how to update a large range, see [Write a large dataset](../resources/samples/write-large-dataset.md).</span></span>

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> <span data-ttu-id="2d5b7-147">`try...catch`L’utilisation à l’intérieur ou autour d’une boucle ralentit votre script.</span><span class="sxs-lookup"><span data-stu-id="2d5b7-147">Using `try...catch` inside or around a loop slows down your script.</span></span> <span data-ttu-id="2d5b7-148">Pour plus d’informations sur les performances, voir [Éviter d’utiliser `try...catch` des blocs.](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)</span><span class="sxs-lookup"><span data-stu-id="2d5b7-148">For more performance information, see [Avoid using `try...catch` blocks](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops).</span></span>

## <a name="see-also"></a><span data-ttu-id="2d5b7-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2d5b7-149">See also</span></span>

- [<span data-ttu-id="2d5b7-150">Dépannage de Office Scripts</span><span class="sxs-lookup"><span data-stu-id="2d5b7-150">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="2d5b7-151">Informations de dépannage pour les Power Automate avec Office scripts</span><span class="sxs-lookup"><span data-stu-id="2d5b7-151">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="2d5b7-152">Limites de plateforme avec Office scripts</span><span class="sxs-lookup"><span data-stu-id="2d5b7-152">Platform limits with Office Scripts</span></span>](../testing/platform-limits.md)
- [<span data-ttu-id="2d5b7-153">Améliorer les performances de vos scripts Office de gestion</span><span class="sxs-lookup"><span data-stu-id="2d5b7-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
