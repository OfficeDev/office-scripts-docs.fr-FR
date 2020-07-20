---
title: Améliorer les performances de vos scripts Office
description: Créez des scripts plus rapides en vous familiarisant avec la communication entre le classeur Excel et votre script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878805"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="3a564-103">Améliorer les performances de vos scripts Office</span><span class="sxs-lookup"><span data-stu-id="3a564-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="3a564-104">L’objectif des scripts Office est d’automatiser la série de tâches couramment exécutées pour vous permettre de gagner du temps.</span><span class="sxs-lookup"><span data-stu-id="3a564-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="3a564-105">Un script lent peut sembler n’accélérer pas votre flux de travail.</span><span class="sxs-lookup"><span data-stu-id="3a564-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="3a564-106">La plupart du temps, votre script sera parfait et s’exécutera comme prévu.</span><span class="sxs-lookup"><span data-stu-id="3a564-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="3a564-107">Toutefois, il existe quelques scénarios évitables qui peuvent affecter les performances.</span><span class="sxs-lookup"><span data-stu-id="3a564-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="3a564-108">La cause la plus fréquente d’un script lent est une communication excessive avec le classeur.</span><span class="sxs-lookup"><span data-stu-id="3a564-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="3a564-109">Votre script s’exécute sur votre ordinateur local, tandis que le classeur existe dans le Cloud.</span><span class="sxs-lookup"><span data-stu-id="3a564-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="3a564-110">À certains moments, votre script synchronise ses données locales avec celles du classeur.</span><span class="sxs-lookup"><span data-stu-id="3a564-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="3a564-111">Cela signifie que toutes les opérations d’écriture (telles que `workbook.addWorksheet()` ) sont appliquées au classeur uniquement lorsque cette synchronisation en arrière-plan se produit.</span><span class="sxs-lookup"><span data-stu-id="3a564-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="3a564-112">De même, toutes les opérations de lecture (telles que `myRange.getValues()` ) obtiennent uniquement les données du classeur pour le script à ces moments.</span><span class="sxs-lookup"><span data-stu-id="3a564-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="3a564-113">Dans les deux cas, le script récupère les informations avant qu’il agisse sur les données.</span><span class="sxs-lookup"><span data-stu-id="3a564-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="3a564-114">Par exemple, le code suivant consigne exactement le nombre de lignes dans la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="3a564-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="3a564-115">API de scripts Office Assurez-vous que toutes les données du classeur ou du script sont exactes et à jour, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="3a564-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="3a564-116">Vous n’avez pas à vous soucier de ces synchronisations pour que votre script s’exécute correctement.</span><span class="sxs-lookup"><span data-stu-id="3a564-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="3a564-117">Toutefois, une connaissance de cette communication de script vers le Cloud peut vous aider à éviter les appels réseau inutiles.</span><span class="sxs-lookup"><span data-stu-id="3a564-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="3a564-118">Optimisation des performances</span><span class="sxs-lookup"><span data-stu-id="3a564-118">Performance optimizations</span></span>

<span data-ttu-id="3a564-119">Vous pouvez appliquer des techniques simples pour réduire la communication vers le Cloud.</span><span class="sxs-lookup"><span data-stu-id="3a564-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="3a564-120">Les modèles suivants permettent d’accélérer vos scripts.</span><span class="sxs-lookup"><span data-stu-id="3a564-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="3a564-121">Lire les données du classeur une seule fois au lieu de répéter dans une boucle.</span><span class="sxs-lookup"><span data-stu-id="3a564-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="3a564-122">Supprimez les `console.log` instructions inutiles.</span><span class="sxs-lookup"><span data-stu-id="3a564-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="3a564-123">Évitez d’utiliser des blocs try/catch.</span><span class="sxs-lookup"><span data-stu-id="3a564-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="3a564-124">Lire les données du classeur en dehors d’une boucle</span><span class="sxs-lookup"><span data-stu-id="3a564-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="3a564-125">Toute méthode qui obtient les données du classeur peut déclencher un appel réseau.</span><span class="sxs-lookup"><span data-stu-id="3a564-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="3a564-126">Au lieu de faire le même appel de manière répétée, vous devez enregistrer les données localement chaque fois que cela est possible.</span><span class="sxs-lookup"><span data-stu-id="3a564-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="3a564-127">Cela est particulièrement vrai pour le traitement des boucles.</span><span class="sxs-lookup"><span data-stu-id="3a564-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="3a564-128">Considérez un script pour obtenir le nombre de nombres négatifs dans la plage utilisée d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="3a564-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="3a564-129">Le script doit parcourir toutes les cellules de la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="3a564-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="3a564-130">Pour ce faire, il a besoin de la plage, du nombre de lignes et du nombre de colonnes.</span><span class="sxs-lookup"><span data-stu-id="3a564-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="3a564-131">Vous devez les stocker en tant que variables locales avant de lancer la boucle.</span><span class="sxs-lookup"><span data-stu-id="3a564-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="3a564-132">Dans le cas contraire, chaque itération de la boucle forcera un retour au classeur.</span><span class="sxs-lookup"><span data-stu-id="3a564-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="3a564-133">À titre d’expérimentation, essayez `usedRangeValues` de remplacer dans la boucle par `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="3a564-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="3a564-134">Vous pouvez remarquer que l’exécution du script est beaucoup plus longue lorsque vous traitez des grandes plages.</span><span class="sxs-lookup"><span data-stu-id="3a564-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="3a564-135">Supprimer les `console.log` instructions inutiles</span><span class="sxs-lookup"><span data-stu-id="3a564-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="3a564-136">La journalisation de console est un outil essentiel pour [le débogage de vos scripts](../testing/troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="3a564-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="3a564-137">Toutefois, il force le script à se synchroniser avec le classeur afin de s’assurer que les informations consignées sont à jour.</span><span class="sxs-lookup"><span data-stu-id="3a564-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="3a564-138">Envisagez de supprimer les instructions de journalisation inutiles (telles que celles utilisées pour les tests) avant de partager votre script.</span><span class="sxs-lookup"><span data-stu-id="3a564-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="3a564-139">Cela ne provoque généralement pas de problèmes de performances perceptibles, sauf si l' `console.log()` instruction est en boucle.</span><span class="sxs-lookup"><span data-stu-id="3a564-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="3a564-140">Éviter d’utiliser des blocs try/catch</span><span class="sxs-lookup"><span data-stu-id="3a564-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="3a564-141">Nous vous déconseillons d’utiliser des [ `try` / `catch` blocs](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) dans le cadre du flux de contrôle attendu d’un script.</span><span class="sxs-lookup"><span data-stu-id="3a564-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="3a564-142">La plupart des erreurs peuvent être évitées en vérifiant les objets renvoyés à partir du classeur.</span><span class="sxs-lookup"><span data-stu-id="3a564-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="3a564-143">Par exemple, le script suivant vérifie que la table renvoyée par le classeur existe avant d’essayer d’ajouter une ligne.</span><span class="sxs-lookup"><span data-stu-id="3a564-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a><span data-ttu-id="3a564-144">Aide cas par cas</span><span class="sxs-lookup"><span data-stu-id="3a564-144">Case-by-case help</span></span>

<span data-ttu-id="3a564-145">À mesure que la plateforme de scripts Office s’étend pour fonctionner avec [Power automate](https://flow.microsoft.com/), [cartes adaptatives](https://docs.microsoft.com/adaptive-cards)et autres fonctionnalités de produit, les détails de la communication de classeur de script deviennent plus compliqués.</span><span class="sxs-lookup"><span data-stu-id="3a564-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="3a564-146">Si vous avez besoin d’aide pour faire en sorte que votre script s’exécute plus rapidement, veuillez contacter le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="3a564-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="3a564-147">N’oubliez pas de baliser votre question avec « Office-script » afin que les experts puissent y trouver des rubriques et de l’aide.</span><span class="sxs-lookup"><span data-stu-id="3a564-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="3a564-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3a564-148">See also</span></span>

- [<span data-ttu-id="3a564-149">Principes de base des scripts pour Office Scripts dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="3a564-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="3a564-150">NOTIFICATION Web docs : boucles et itération</span><span class="sxs-lookup"><span data-stu-id="3a564-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
