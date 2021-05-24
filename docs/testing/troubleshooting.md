---
title: Résoudre les problèmes Office scripts
description: Conseils et techniques de débogage pour Office scripts, ainsi que des ressources d’aide.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545554"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="7adc4-103">Résoudre les problèmes Office scripts</span><span class="sxs-lookup"><span data-stu-id="7adc4-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="7adc4-104">Lorsque vous développez Office scripts, vous pouvez faire des erreurs.</span><span class="sxs-lookup"><span data-stu-id="7adc4-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="7adc4-105">C'est bon.</span><span class="sxs-lookup"><span data-stu-id="7adc4-105">It's okay.</span></span> <span data-ttu-id="7adc4-106">Vous avez les outils nécessaires pour trouver les problèmes et faire fonctionner parfaitement vos scripts.</span><span class="sxs-lookup"><span data-stu-id="7adc4-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="7adc4-107">Types d’erreurs</span><span class="sxs-lookup"><span data-stu-id="7adc4-107">Types of errors</span></span>

<span data-ttu-id="7adc4-108">Office Les erreurs de script se classent dans l’une des deux catégories suivantes :</span><span class="sxs-lookup"><span data-stu-id="7adc4-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="7adc4-109">Erreurs ou avertissements au moment de la compilation</span><span class="sxs-lookup"><span data-stu-id="7adc4-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="7adc4-110">Erreurs d’runtime</span><span class="sxs-lookup"><span data-stu-id="7adc4-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="7adc4-111">Erreurs au moment de la compilation</span><span class="sxs-lookup"><span data-stu-id="7adc4-111">Compile-time errors</span></span>

<span data-ttu-id="7adc4-112">Les erreurs et avertissements au moment de la compilation sont initialement affichés dans l’Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="7adc4-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="7adc4-113">Ces éléments sont affichés par les soulignements ondulés rouges dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="7adc4-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="7adc4-114">Ils sont également affichés sous l’onglet **Problèmes** en bas du volet Des tâches de l’Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="7adc4-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="7adc4-115">La sélection de l’erreur donne plus de détails sur le problème et suggère des solutions.</span><span class="sxs-lookup"><span data-stu-id="7adc4-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="7adc4-116">Les erreurs de compilation doivent être traitées avant l’exécution du script.</span><span class="sxs-lookup"><span data-stu-id="7adc4-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Une erreur de compilateur affichée dans le texte de pointeur de l’éditeur de code":::

<span data-ttu-id="7adc4-118">Vous pouvez également voir des soulignements d’avertissement orange et des messages d’information gris.</span><span class="sxs-lookup"><span data-stu-id="7adc4-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="7adc4-119">Celles-ci indiquent des suggestions de performances ou d’autres possibilités dans le cas où le script peut avoir des effets involontaires.</span><span class="sxs-lookup"><span data-stu-id="7adc4-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="7adc4-120">Ces avertissements doivent être examinés attentivement avant de les ignorer.</span><span class="sxs-lookup"><span data-stu-id="7adc4-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="7adc4-121">Erreurs d’runtime</span><span class="sxs-lookup"><span data-stu-id="7adc4-121">Runtime errors</span></span>

<span data-ttu-id="7adc4-122">Les erreurs d’utilisation se produisent en raison de problèmes logiques dans le script.</span><span class="sxs-lookup"><span data-stu-id="7adc4-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="7adc4-123">Cela peut être dû au fait qu’un objet utilisé dans le script ne se trouve pas dans le workbook, qu’un tableau est formaté différemment des prévisions ou qu’il existe une légère différence entre les exigences du script et le workbook actuel.</span><span class="sxs-lookup"><span data-stu-id="7adc4-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="7adc4-124">Le script suivant génère une erreur lorsqu’une feuille de calcul nommée « TestSheet » n’est pas présente.</span><span class="sxs-lookup"><span data-stu-id="7adc4-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="7adc4-125">Messages de la console</span><span class="sxs-lookup"><span data-stu-id="7adc4-125">Console messages</span></span>

<span data-ttu-id="7adc4-126">Les erreurs de compilation et d’runtime affichent des messages d’erreur dans la console lorsqu’un script s’exécute.</span><span class="sxs-lookup"><span data-stu-id="7adc4-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="7adc4-127">Ils donnent un numéro de ligne où le problème s’est produits.</span><span class="sxs-lookup"><span data-stu-id="7adc4-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="7adc4-128">N’oubliez pas que la cause première d’un problème peut être une ligne de code différente de ce qui est indiqué dans la console.</span><span class="sxs-lookup"><span data-stu-id="7adc4-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="7adc4-129">L’image suivante montre la sortie de la console pour [l’erreur `any` ](../develop/typescript-restrictions.md) explicite du compilateur.</span><span class="sxs-lookup"><span data-stu-id="7adc4-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="7adc4-130">Notez le texte `[5, 16]` au début de la chaîne d’erreur.</span><span class="sxs-lookup"><span data-stu-id="7adc4-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="7adc4-131">Cela indique que l’erreur se trouve sur la ligne 5, en commençant au caractère 16.</span><span class="sxs-lookup"><span data-stu-id="7adc4-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La console Éditeur de code affichant un message d’erreur explicite « tout »":::

<span data-ttu-id="7adc4-133">L’image suivante montre la sortie de la console pour une erreur d’runtime.</span><span class="sxs-lookup"><span data-stu-id="7adc4-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="7adc4-134">Ici, le script tente d’ajouter une feuille de calcul avec le nom d’une feuille de calcul existante.</span><span class="sxs-lookup"><span data-stu-id="7adc4-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="7adc4-135">Là encore, notez la « ligne 2 » précédant l’erreur pour afficher la ligne à examiner.</span><span class="sxs-lookup"><span data-stu-id="7adc4-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="La console Éditeur de code affichant une erreur à partir de l’appel « addWorksheet »":::

## <a name="console-logs"></a><span data-ttu-id="7adc4-137">Journaux de la console</span><span class="sxs-lookup"><span data-stu-id="7adc4-137">Console logs</span></span>

<span data-ttu-id="7adc4-138">Imprime les messages à l’écran avec `console.log` l’instruction.</span><span class="sxs-lookup"><span data-stu-id="7adc4-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="7adc4-139">Ces journaux peuvent vous montrer la valeur actuelle des variables ou les chemins de code qui sont déclenchés.</span><span class="sxs-lookup"><span data-stu-id="7adc4-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="7adc4-140">Pour ce faire, `console.log` appelez avec n’importe quel objet en tant que paramètre.</span><span class="sxs-lookup"><span data-stu-id="7adc4-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="7adc4-141">En règle générale, `string` il s’agit du type le plus simple à lire dans la console.</span><span class="sxs-lookup"><span data-stu-id="7adc4-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="7adc4-142">Les chaînes transmises sont affichées dans la console de journalisation de l’éditeur de code, en `console.log` bas du volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="7adc4-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="7adc4-143">Les journaux se  trouvent sous l’onglet Sortie, bien que l’onglet soit automatiquement mis au point lors de l’écriture d’un journal.</span><span class="sxs-lookup"><span data-stu-id="7adc4-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="7adc4-144">Les journaux n’affectent pas le workbook.</span><span class="sxs-lookup"><span data-stu-id="7adc4-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="7adc4-145">Automatiser l’onglet qui n’apparaît pas ou Office scripts indisponibles</span><span class="sxs-lookup"><span data-stu-id="7adc4-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="7adc4-146">Les étapes suivantes doivent vous aider à résoudre les problèmes liés à l’onglet **Automatiser** qui n’apparaît pas dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="7adc4-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="7adc4-147">[Assurez-vous que votre licence Microsoft 365 inclut Office scripts.](../overview/excel.md#requirements)</span><span class="sxs-lookup"><span data-stu-id="7adc4-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="7adc4-148">[Vérifiez que votre navigateur est pris en charge.](platform-limits.md#browser-support)</span><span class="sxs-lookup"><span data-stu-id="7adc4-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="7adc4-149">[Assurez-vous que les cookies tiers sont activés.](platform-limits.md#third-party-cookies)</span><span class="sxs-lookup"><span data-stu-id="7adc4-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="7adc4-150">[Assurez-vous que votre administrateur n’a pas désactivé Office scripts dans le centre Microsoft 365'administration.](/microsoft-365/admin/manage/manage-office-scripts-settings)</span><span class="sxs-lookup"><span data-stu-id="7adc4-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="7adc4-151">Résoudre les problèmes de scripts dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="7adc4-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="7adc4-152">Pour plus d’informations sur l’exécution de scripts Power Automate, voir Résolution des problèmes Office [scripts en](power-automate-troubleshooting.md)cours d’exécution dans Power Automate .</span><span class="sxs-lookup"><span data-stu-id="7adc4-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="7adc4-153">Ressources d’aide</span><span class="sxs-lookup"><span data-stu-id="7adc4-153">Help resources</span></span>

<span data-ttu-id="7adc4-154">[Stack Overflow est](https://stackoverflow.com/questions/tagged/office-scripts) une communauté de développeurs prêts à vous aider avec les problèmes de codage.</span><span class="sxs-lookup"><span data-stu-id="7adc4-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="7adc4-155">Souvent, vous serez en mesure de trouver la solution à votre problème par le biais d’une recherche rapide de stack overflow.</span><span class="sxs-lookup"><span data-stu-id="7adc4-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="7adc4-156">Si ce n’est pas le cas, posez votre question et marquez-la avec la balise « office-scripts ».</span><span class="sxs-lookup"><span data-stu-id="7adc4-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="7adc4-157">N’oubliez pas de mentionner que vous créez un *script* Office, et non un *Office.*</span><span class="sxs-lookup"><span data-stu-id="7adc4-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="7adc4-158">Si vous rencontrez un problème avec l’API JavaScript Office, créez un problème dans le référentiel [officeDev/office-js](https://github.com/OfficeDev/office-js) GitHub.</span><span class="sxs-lookup"><span data-stu-id="7adc4-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="7adc4-159">Les membres de l’équipe produit répondent aux problèmes et fournissent une assistance supplémentaire.</span><span class="sxs-lookup"><span data-stu-id="7adc4-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="7adc4-160">La création d’un problème dans le référentiel **OfficeDev/office-js** indique que vous avez trouvé une faille dans la bibliothèque d’API JavaScript Office que l’équipe du produit doit résoudre.</span><span class="sxs-lookup"><span data-stu-id="7adc4-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="7adc4-161">En cas de problème avec l’enregistreur d’actions ou l’éditeur, envoyez des commentaires via le bouton **d'> commentaires** dans Excel.</span><span class="sxs-lookup"><span data-stu-id="7adc4-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="7adc4-162">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7adc4-162">See also</span></span>

- [<span data-ttu-id="7adc4-163">Meilleures pratiques dans Office scripts</span><span class="sxs-lookup"><span data-stu-id="7adc4-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="7adc4-164">Limites de plateforme avec Office scripts</span><span class="sxs-lookup"><span data-stu-id="7adc4-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="7adc4-165">Améliorer les performances de vos scripts Office de gestion</span><span class="sxs-lookup"><span data-stu-id="7adc4-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="7adc4-166">Résoudre les Office scripts en cours d’exécution dans PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="7adc4-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="7adc4-167">Annuler les effets des scripts Office scripts</span><span class="sxs-lookup"><span data-stu-id="7adc4-167">Undo the effects of Office Scripts</span></span>](undo.md)
