---
title: Améliorez les performances de vos scripts Office’argent
description: Créez des scripts plus rapides en comprenant la communication entre Excel manuel et votre script.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 512e2108cb81cf9ac8ae98980951d5d01b3d2de9
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52544990"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="0200d-103">Améliorez les performances de vos scripts Office’argent</span><span class="sxs-lookup"><span data-stu-id="0200d-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="0200d-104">Le but de Office scripts est d’automatiser des séries de tâches couramment exécutées pour vous faire gagner du temps.</span><span class="sxs-lookup"><span data-stu-id="0200d-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="0200d-105">Un script lent peut avoir l’impression qu’il n’accélère pas votre flux de travail.</span><span class="sxs-lookup"><span data-stu-id="0200d-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="0200d-106">La plupart du temps, votre script sera parfaitement bien et exécuté comme prévu.</span><span class="sxs-lookup"><span data-stu-id="0200d-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="0200d-107">Cependant, il existe quelques scénarios évitables qui peuvent affecter les performances.</span><span class="sxs-lookup"><span data-stu-id="0200d-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="0200d-108">La raison la plus courante d’un script lent est une communication excessive avec le cahier de travail.</span><span class="sxs-lookup"><span data-stu-id="0200d-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="0200d-109">Votre script s’exécute sur votre machine locale, tandis que le manuel existe dans le cloud.</span><span class="sxs-lookup"><span data-stu-id="0200d-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="0200d-110">À certains moments, votre script synchronise ses données locales avec celle du cahier de travail.</span><span class="sxs-lookup"><span data-stu-id="0200d-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="0200d-111">Cela signifie que toutes les opérations d’écriture (telles `workbook.addWorksheet()` que) ne sont appliquées au manuel que lorsque cette synchronisation en coulisses se produit.</span><span class="sxs-lookup"><span data-stu-id="0200d-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="0200d-112">De même, toutes les opérations de lecture (telles `myRange.getValues()` que) ne proviennent que des données du cahier de travail pour le script à ces moments.</span><span class="sxs-lookup"><span data-stu-id="0200d-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="0200d-113">Dans les deux cas, le script récupère des informations avant qu’elles n’agissent sur les données.</span><span class="sxs-lookup"><span data-stu-id="0200d-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="0200d-114">Par exemple, le code suivant enregistrera avec précision le nombre de lignes dans la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="0200d-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="0200d-115">Office Les API scripts garantissent que toutes les données du cahier de travail ou du script sont exactes et à jour si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="0200d-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="0200d-116">Vous n’avez pas besoin de vous soucier de ces synchronisations pour que votre script s’exécute correctement.</span><span class="sxs-lookup"><span data-stu-id="0200d-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="0200d-117">Toutefois, une prise de conscience de cette communication script-cloud peut vous aider à éviter les appels réseau non désaillés.</span><span class="sxs-lookup"><span data-stu-id="0200d-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="0200d-118">Optimisations de performance</span><span class="sxs-lookup"><span data-stu-id="0200d-118">Performance optimizations</span></span>

<span data-ttu-id="0200d-119">Vous pouvez appliquer des techniques simples pour aider à réduire la communication vers le cloud.</span><span class="sxs-lookup"><span data-stu-id="0200d-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="0200d-120">Les modèles suivants aident à accélérer vos scripts.</span><span class="sxs-lookup"><span data-stu-id="0200d-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="0200d-121">Lisez les données du manuel une fois au lieu de plusieurs fois en boucle.</span><span class="sxs-lookup"><span data-stu-id="0200d-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="0200d-122">Supprimez les `console.log` instructions inutiles.</span><span class="sxs-lookup"><span data-stu-id="0200d-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="0200d-123">Évitez d’utiliser des blocs try/catch.</span><span class="sxs-lookup"><span data-stu-id="0200d-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="0200d-124">Lire les données du cahier de travail en dehors d’une boucle</span><span class="sxs-lookup"><span data-stu-id="0200d-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="0200d-125">Toute méthode qui obtient des données du cahier de travail peut déclencher un appel réseau.</span><span class="sxs-lookup"><span data-stu-id="0200d-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="0200d-126">Plutôt que de faire à plusieurs reprises le même appel, vous devez enregistrer des données localement dans la mesure du possible.</span><span class="sxs-lookup"><span data-stu-id="0200d-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="0200d-127">Cela est particulièrement vrai lorsqu’il s’agit de boucles.</span><span class="sxs-lookup"><span data-stu-id="0200d-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="0200d-128">Considérez un script pour obtenir le nombre de nombres négatifs dans la plage utilisée d’une feuille de travail.</span><span class="sxs-lookup"><span data-stu-id="0200d-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="0200d-129">Le script doit itérer sur chaque cellule de la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="0200d-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="0200d-130">Pour ce faire, il a besoin de la plage, le nombre de lignes, et le nombre de colonnes.</span><span class="sxs-lookup"><span data-stu-id="0200d-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="0200d-131">Vous devez les stocker comme variables locales avant de commencer la boucle.</span><span class="sxs-lookup"><span data-stu-id="0200d-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="0200d-132">Dans le cas contraire, chaque itération de la boucle forcera un retour au cahier de travail.</span><span class="sxs-lookup"><span data-stu-id="0200d-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

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
> <span data-ttu-id="0200d-133">Comme une expérience, essayez de remplacer `usedRangeValues` dans la boucle par `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="0200d-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="0200d-134">Vous remarquerez peut-être que le script prend beaucoup plus de temps à exécuter lorsqu’il s’agit de grandes plages.</span><span class="sxs-lookup"><span data-stu-id="0200d-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a><span data-ttu-id="0200d-135">Évitez d’utiliser `try...catch` des blocs dans ou autour des boucles</span><span class="sxs-lookup"><span data-stu-id="0200d-135">Avoid using `try...catch` blocks in or surrounding loops</span></span>

<span data-ttu-id="0200d-136">Nous ne recommandons pas d’utiliser les [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) instructions en boucle ou en boucles environnantes.</span><span class="sxs-lookup"><span data-stu-id="0200d-136">We don't recommend using [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) statements either in loops or surrounding loops.</span></span> <span data-ttu-id="0200d-137">C’est pour la même raison que vous devez éviter de lire les données en boucle : chaque itération oblige le script à se synchroniser avec le cahier de travail pour s’assurer qu’aucune erreur n’a été lancée.</span><span class="sxs-lookup"><span data-stu-id="0200d-137">This is for the same reason you should avoid reading data in a loop: each iteration forces the script to synchronize with the workbook to make sure no error has been thrown.</span></span> <span data-ttu-id="0200d-138">La plupart des erreurs peuvent être évitées en vérifiant les objets retournés du cahier de travail.</span><span class="sxs-lookup"><span data-stu-id="0200d-138">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="0200d-139">Par exemple, le script suivant vérifie que la table retournée par le cahier de travail existe avant d’essayer d’ajouter une ligne.</span><span class="sxs-lookup"><span data-stu-id="0200d-139">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

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

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="0200d-140">Supprimer les instructions `console.log` inutiles</span><span class="sxs-lookup"><span data-stu-id="0200d-140">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="0200d-141">L’enregistrement des consoles est un outil essentiel [pour débouger vos scripts.](../testing/troubleshooting.md)</span><span class="sxs-lookup"><span data-stu-id="0200d-141">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="0200d-142">Toutefois, il oblige le script à se synchroniser avec le cahier de travail pour s’assurer que les informations enregistrées sont à jour.</span><span class="sxs-lookup"><span data-stu-id="0200d-142">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="0200d-143">Envisagez de supprimer les instructions d’enregistrement inutiles (telles que celles utilisées pour les tests) avant de partager votre script.</span><span class="sxs-lookup"><span data-stu-id="0200d-143">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="0200d-144">Cela ne cause généralement pas de problème de performances notables, sauf si `console.log()` l’instruction est en boucle.</span><span class="sxs-lookup"><span data-stu-id="0200d-144">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

## <a name="case-by-case-help"></a><span data-ttu-id="0200d-145">Aide au cas par cas</span><span class="sxs-lookup"><span data-stu-id="0200d-145">Case-by-case help</span></span>

<span data-ttu-id="0200d-146">Au fur et à mesure que la plate-forme Office Scripts [s’étend pour fonctionner avec Power Automate,](https://flow.microsoft.com/)Adaptive [Cards](/adaptive-cards)et d’autres fonctionnalités de produits croisés, les détails de la communication script-cahier de travail deviennent plus complexes.</span><span class="sxs-lookup"><span data-stu-id="0200d-146">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="0200d-147">Si vous avez besoin d’aide pour faire fonctionner votre script plus rapidement, s’il vous plaît [tendre la main via Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span><span class="sxs-lookup"><span data-stu-id="0200d-147">If you need help making your script run faster, please reach out through [Microsoft Q&A](/answers/topics/office-scripts-dev.html).</span></span> <span data-ttu-id="0200d-148">Assurez-vous d’étiqueter votre question avec « office-scripts-dev » afin que les experts puissent la trouver et vous aider.</span><span class="sxs-lookup"><span data-stu-id="0200d-148">Be sure to tag your question with "office-scripts-dev" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="0200d-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0200d-149">See also</span></span>

- [<span data-ttu-id="0200d-150">Principes de base pour la rédaction de scripts Office en Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="0200d-150">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="0200d-151">MDN web docs: Boucles et itération</span><span class="sxs-lookup"><span data-stu-id="0200d-151">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
