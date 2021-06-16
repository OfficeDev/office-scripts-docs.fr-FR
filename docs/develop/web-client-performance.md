---
title: Améliorer les performances de vos scripts Office de gestion
description: Créez des scripts plus rapides en comprenant la communication entre le Excel et votre script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: a5bd879625b9c3bac0caa621dde312f7c961dd5c
ms.sourcegitcommit: 2aaf7dc527cb6c9f1206550b2c5745280503b2a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/16/2021
ms.locfileid: "52957699"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="3c760-103">Améliorer les performances de vos scripts Office de gestion</span><span class="sxs-lookup"><span data-stu-id="3c760-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="3c760-104">L’objectif de Office Scripts est d’automatiser une série de tâches couramment exécutées pour vous faire gagner du temps.</span><span class="sxs-lookup"><span data-stu-id="3c760-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="3c760-105">Un script lent peut avoir l’impression qu’il n’accélère pas votre flux de travail.</span><span class="sxs-lookup"><span data-stu-id="3c760-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="3c760-106">La plupart du temps, votre script sera parfaitement correct et s’exécutera comme prévu.</span><span class="sxs-lookup"><span data-stu-id="3c760-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="3c760-107">Toutefois, il existe quelques scénarios qui peuvent avoir une incidence sur les performances.</span><span class="sxs-lookup"><span data-stu-id="3c760-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="3c760-108">La raison la plus courante d’un script lent est une communication excessive avec le workbook.</span><span class="sxs-lookup"><span data-stu-id="3c760-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="3c760-109">Votre script s’exécute sur votre ordinateur local, tandis que le workbook existe dans le cloud.</span><span class="sxs-lookup"><span data-stu-id="3c760-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="3c760-110">À certains moments, votre script synchronise ses données locales avec celle du workbook.</span><span class="sxs-lookup"><span data-stu-id="3c760-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="3c760-111">Cela signifie que toutes les opérations d’écriture (telles que ) ne sont appliquées aubook que lorsque cette synchronisation en `workbook.addWorksheet()` arrière-plan se produit.</span><span class="sxs-lookup"><span data-stu-id="3c760-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="3c760-112">De même, toutes les opérations de lecture (telles que ) obtiennent uniquement des données du manuel pour le `myRange.getValues()` script à ce moment-là.</span><span class="sxs-lookup"><span data-stu-id="3c760-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="3c760-113">Dans les deux cas, le script récupère des informations avant d’agir sur les données.</span><span class="sxs-lookup"><span data-stu-id="3c760-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="3c760-114">Par exemple, le code suivant enregistre avec précision le nombre de lignes dans la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="3c760-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="3c760-115">Office Les API de scripts garantissent que toutes les données du workbook ou du script sont précises et à jour si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="3c760-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="3c760-116">Vous n’avez pas besoin de vous soucier de ces synchronisations pour que votre script s’exécute correctement.</span><span class="sxs-lookup"><span data-stu-id="3c760-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="3c760-117">Toutefois, une connaissance de cette communication entre les scripts et le cloud peut vous aider à éviter les appels réseau inutiles.</span><span class="sxs-lookup"><span data-stu-id="3c760-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="3c760-118">Optimisation des performances</span><span class="sxs-lookup"><span data-stu-id="3c760-118">Performance optimizations</span></span>

<span data-ttu-id="3c760-119">Vous pouvez appliquer des techniques simples pour réduire la communication vers le cloud.</span><span class="sxs-lookup"><span data-stu-id="3c760-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="3c760-120">Les modèles suivants permettent d’accélérer vos scripts.</span><span class="sxs-lookup"><span data-stu-id="3c760-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="3c760-121">Lire les données du workbook une seule fois plutôt que de manière répétée dans une boucle.</span><span class="sxs-lookup"><span data-stu-id="3c760-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="3c760-122">Supprimez les `console.log` instructions inutiles.</span><span class="sxs-lookup"><span data-stu-id="3c760-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="3c760-123">Évitez d’utiliser des blocs try/catch.</span><span class="sxs-lookup"><span data-stu-id="3c760-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="3c760-124">Lire les données d’un workbook en dehors d’une boucle</span><span class="sxs-lookup"><span data-stu-id="3c760-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="3c760-125">Toute méthode qui obtient des données à partir du manuel peut déclencher un appel réseau.</span><span class="sxs-lookup"><span data-stu-id="3c760-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="3c760-126">Au lieu d’effectuer le même appel à plusieurs reprises, vous devez enregistrer les données localement chaque fois que cela est possible.</span><span class="sxs-lookup"><span data-stu-id="3c760-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="3c760-127">Cela est particulièrement vrai lorsque vous traitez des boucles.</span><span class="sxs-lookup"><span data-stu-id="3c760-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="3c760-128">Envisagez un script pour obtenir le nombre de nombres négatifs dans la plage utilisée d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="3c760-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="3c760-129">Le script doit itérer sur chaque cellule de la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="3c760-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="3c760-130">Pour ce faire, il a besoin de la plage, du nombre de lignes et du nombre de colonnes.</span><span class="sxs-lookup"><span data-stu-id="3c760-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="3c760-131">Vous devez les stocker en tant que variables locales avant de démarrer la boucle.</span><span class="sxs-lookup"><span data-stu-id="3c760-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="3c760-132">Dans le cas contraire, chaque itération de la boucle force un retour au workbook.</span><span class="sxs-lookup"><span data-stu-id="3c760-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="3c760-133">En tant qu’expérience, essayez de remplacer `usedRangeValues` dans la boucle par `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="3c760-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="3c760-134">Vous remarquerez peut-être que l’exécuter du script prend beaucoup plus de temps lorsque vous traitez des plages importantes.</span><span class="sxs-lookup"><span data-stu-id="3c760-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="3c760-135">Évitez `try...catch` d’utiliser des blocs dans des boucles dans ou autour</span><span class="sxs-lookup"><span data-stu-id="3c760-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="3c760-136">Nous vous déconseillons d’utiliser des instructions en [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) boucle ou autour de boucles.</span><span class="sxs-lookup"><span data-stu-id="3c760-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="3c760-137">C’est pour la même raison que vous devez éviter la lecture de données dans une boucle : chaque itération force le script à se synchroniser avec le workbook pour s’assurer qu’aucune erreur n’a été lancée.</span><span class="sxs-lookup"><span data-stu-id="3c760-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="3c760-138">La plupart des erreurs peuvent être évitées en vérifiant les objets renvoyés à partir du workbook.</span><span class="sxs-lookup"><span data-stu-id="3c760-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="3c760-139">Par exemple, le script suivant vérifie que la table renvoyée par le workbook existe avant d’essayer d’ajouter une ligne.</span><span class="sxs-lookup"><span data-stu-id="3c760-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="3c760-140">Supprimer les `console.log` instructions inutiles</span><span class="sxs-lookup"><span data-stu-id="3c760-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="3c760-141">La journalisation de la console est un outil vital [pour le débogage de vos scripts.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="3c760-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="3c760-142">Toutefois, il force le script à se synchroniser avec le workbook pour s’assurer que les informations consignées sont à jour.</span><span class="sxs-lookup"><span data-stu-id="3c760-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="3c760-143">Envisagez de supprimer les instructions de journalisation inutiles (telles que celles utilisées pour les tests) avant de partager votre script.</span><span class="sxs-lookup"><span data-stu-id="3c760-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="3c760-144">En règle générale, cela ne provoque pas de problème de performances perceptible, sauf si `console.log()` l’instruction est en boucle.</span><span class="sxs-lookup"><span data-stu-id="3c760-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="3c760-145">Aide au cas par cas</span><span class="sxs-lookup"><span data-stu-id="3c760-145">Case-by-case help</span></span>

<span data-ttu-id="3c760-146">À mesure que la plateforme Office Scripts s’étend [](/adaptive-cards)pour fonctionner avec [Power Automate,](https://flow.microsoft.com/)les cartes adaptatives et d’autres fonctionnalités entre produits, les détails de la communication de script-workbook deviennent plus complexes.</span><span class="sxs-lookup"><span data-stu-id="3c760-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="3c760-147">Si vous avez besoin d’aide pour accélérer l’exécuter, [contactez-vous](/answers/topics/office-scripts-excel-dev.html)via Microsoft Q&A .</span><span class="sxs-lookup"><span data-stu-id="3c760-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-excel-dev.html).</span></span> <span data-ttu-id="3c760-148">N’oubliez pas de baliser votre question avec « office-scripts-dev » afin que les experts la trouvent et vous aident.</span><span class="sxs-lookup"><span data-stu-id="3c760-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="3c760-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3c760-149">See also</span></span>

- [<span data-ttu-id="3c760-150">Principes de base pour la rédaction de scripts Office en Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="3c760-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="3c760-151">Documentation web MDN : boucles et itération</span><span class="sxs-lookup"><span data-stu-id="3c760-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
