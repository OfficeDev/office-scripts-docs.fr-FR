---
title: Scripts de Office dépannage
description: Débogage des conseils et des techniques pour Office scripts, ainsi que des ressources d’aide.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: ff0ac1e63084c7c541d2a4925f1f011d16fa4992
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545554"
---
# <a name="troubleshoot-office-scripts"></a><span data-ttu-id="a6b7d-103">Scripts de Office dépannage</span><span class="sxs-lookup"><span data-stu-id="a6b7d-103">Troubleshoot Office Scripts</span></span>

<span data-ttu-id="a6b7d-104">Au fur et à mesure Office scripts, vous pouvez faire des erreurs.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="a6b7d-105">C'est bon.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-105">It's okay.</span></span> <span data-ttu-id="a6b7d-106">Vous avez les outils pour aider à trouver les problèmes et obtenir vos scripts fonctionnent parfaitement.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-106">You have the tools to help find the problems and get your scripts working perfectly.</span></span>

## <a name="types-of-errors"></a><span data-ttu-id="a6b7d-107">Types d’erreurs</span><span class="sxs-lookup"><span data-stu-id="a6b7d-107">Types of errors</span></span>

<span data-ttu-id="a6b7d-108">Office Les erreurs de script s’insurdent dans l’une des deux catégories suivantes :</span><span class="sxs-lookup"><span data-stu-id="a6b7d-108">Office Scripts errors fall into one of two categories:</span></span>

* <span data-ttu-id="a6b7d-109">Compiler les erreurs ou les avertissements</span><span class="sxs-lookup"><span data-stu-id="a6b7d-109">Compile-time errors or warnings</span></span>
* <span data-ttu-id="a6b7d-110">Erreurs de temps d’exécution</span><span class="sxs-lookup"><span data-stu-id="a6b7d-110">Runtime errors</span></span>

### <a name="compile-time-errors"></a><span data-ttu-id="a6b7d-111">Erreurs de compilement</span><span class="sxs-lookup"><span data-stu-id="a6b7d-111">Compile-time errors</span></span>

<span data-ttu-id="a6b7d-112">Les erreurs et avertissements de compilation sont d’abord affichés dans l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-112">Compile-time errors and warnings are initially shown in the Code Editor.</span></span> <span data-ttu-id="a6b7d-113">Ceux-ci sont montrés par les soulignements rouges ondulés dans l’éditeur.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-113">These are shown by the wavy red underlines in the editor.</span></span> <span data-ttu-id="a6b7d-114">Ils sont également affichés sous **l’onglet Problèmes** au bas du volet de tâche de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-114">They are also displayed under the **Problems** tab at the bottom of the Code Editor task pane.</span></span> <span data-ttu-id="a6b7d-115">Le choix de l’erreur donnera plus de détails sur le problème et proposera des solutions.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-115">Selecting the error will give more details about the problem and suggest solutions.</span></span> <span data-ttu-id="a6b7d-116">Les erreurs de temps de compilation doivent être traitées avant d’exécuter le script.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-116">Compile-time errors should be addressed before running the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="Erreur de compilateur affichée dans le texte stationnaire de l’éditeur de code":::

<span data-ttu-id="a6b7d-118">Vous pouvez également voir des soulignements d’avertissement orange et des messages d’information gris.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-118">You may also see orange warning underlines and grey informational messages.</span></span> <span data-ttu-id="a6b7d-119">Ceux-ci indiquent des suggestions de performances ou d’autres possibilités où le script peut avoir des effets involontaires.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-119">These indicate performance suggestions or other possibilities where the script may have unintentional effects.</span></span> <span data-ttu-id="a6b7d-120">Ces avertissements devraient être examinés de près avant de les rejeter.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-120">Such warnings should be examined closely before dismissing them.</span></span>

### <a name="runtime-errors"></a><span data-ttu-id="a6b7d-121">Erreurs de temps d’exécution</span><span class="sxs-lookup"><span data-stu-id="a6b7d-121">Runtime errors</span></span>

<span data-ttu-id="a6b7d-122">Les erreurs d’exécution se produisent en raison de problèmes logiques dans le script.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-122">Runtime errors happen because of logic issues in the script.</span></span> <span data-ttu-id="a6b7d-123">Cela peut être parce qu’un objet utilisé dans le script n’est pas dans le cahier de travail, une table est formatée différemment que prévu, ou un autre léger écart entre les exigences du script et le manuel actuel.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-123">This could be because an object used in the script isn't in the workbook, a table is formatted differently than anticipated, or some other slight discrepancy between the script's requirements and the current workbook.</span></span> <span data-ttu-id="a6b7d-124">Le script suivant génère une erreur lorsqu’une feuille de travail nommée « Feuille de test » n’est pas présente.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-124">The following script generates an error when a worksheet named "TestSheet" is not present.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let mySheet = workbook.getWorksheet('TestSheet');

  // This will throw an error if there is no "TestSheet".
  mySheet.getRange("A1");
}
```

### <a name="console-messages"></a><span data-ttu-id="a6b7d-125">Messages console</span><span class="sxs-lookup"><span data-stu-id="a6b7d-125">Console messages</span></span>

<span data-ttu-id="a6b7d-126">Les erreurs de compilement et d’exécution affichent les messages d’erreur dans la console lorsqu’un script s’exécute.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-126">Both compile-time and runtime errors display error messages in the console when a script runs.</span></span> <span data-ttu-id="a6b7d-127">Ils donnent un numéro de ligne où le problème a été rencontré.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-127">They give a line number where the problem was encountered.</span></span> <span data-ttu-id="a6b7d-128">Gardez à l’esprit que la cause profonde de tout problème peut être une ligne de code différente de ce qui est indiqué dans la console.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-128">Keep in mind that the root cause of any issue may be a different line of code than what is indicated in the console.</span></span>

<span data-ttu-id="a6b7d-129">L’image suivante affiche la sortie de la console pour [l’erreur `any` compilateur](../develop/typescript-restrictions.md) explicite.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-129">The following image shows the console output for the [explicit `any`](../develop/typescript-restrictions.md) compiler error.</span></span> <span data-ttu-id="a6b7d-130">Notez le texte `[5, 16]` au début de la chaîne d’erreur.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-130">Note the text `[5, 16]` at the beginning of the error string.</span></span> <span data-ttu-id="a6b7d-131">Cela indique que l’erreur est sur la ligne 5, à partir du caractère 16.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-131">This indicates the error is on line 5, starting at character 16.</span></span>
:::image type="content" source="../images/explicit-any-error-message.png" alt-text="La console Code Editor affichant un message d’erreur explicite « n’importe quel »":::

<span data-ttu-id="a6b7d-133">L’image suivante affiche la sortie de la console pour une erreur de temps d’exécution.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-133">The follow image shows the console output for a runtime error.</span></span> <span data-ttu-id="a6b7d-134">Ici, le script tente d’ajouter une feuille de travail avec le nom d’une feuille de travail existante.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-134">Here, the script tries to add a worksheet with a the name of an existing worksheet.</span></span> <span data-ttu-id="a6b7d-135">Encore une fois, notez la « ligne 2 » précédant l’erreur pour montrer quelle ligne enquêter.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-135">Again, note the "Line 2" preceding the error to show which line to investigate.</span></span>
:::image type="content" source="../images/runtime-error-console.png" alt-text="La console Code Editor affichant une erreur de l’appel 'addWorksheet'":::

## <a name="console-logs"></a><span data-ttu-id="a6b7d-137">Journaux de console</span><span class="sxs-lookup"><span data-stu-id="a6b7d-137">Console logs</span></span>

<span data-ttu-id="a6b7d-138">Imprimez des messages à l’écran avec `console.log` l’instruction.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-138">Print messages to the screen with the `console.log` statement.</span></span> <span data-ttu-id="a6b7d-139">Ces journaux peuvent vous montrer la valeur actuelle des variables ou les chemins de code qui sont déclenchés.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-139">These logs can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="a6b7d-140">Pour ce faire, appelez avec `console.log` n’importe quel objet comme paramètre.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-140">To do this, call `console.log` with any object as a parameter.</span></span> <span data-ttu-id="a6b7d-141">Habituellement, un `string` est le type le plus facile à lire dans la console.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-141">Usually, a `string` is the easiest type to read in the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="a6b7d-142">Les chaînes `console.log` transmises sont affichées dans la console de journalisation de l’éditeur de code, au bas du volet de tâche.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-142">Strings passed to `console.log` are displayed in the Code Editor's logging console, at the bottom of the task pane.</span></span> <span data-ttu-id="a6b7d-143">Les journaux se trouvent sur **l’onglet Sortie,** bien que l’onglet gagne automatiquement la mise au point lorsqu’un journal est écrit.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-143">Logs are found on the **Output** tab, though the tab automatically gains focus when a log is written.</span></span>

<span data-ttu-id="a6b7d-144">Les journaux n’affectent pas le cahier de travail.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-144">Logs do not affect the workbook.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="a6b7d-145">Automatisez l’onglet n’apparaissant pas ou n’Office scripts non disponibles</span><span class="sxs-lookup"><span data-stu-id="a6b7d-145">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="a6b7d-146">Les étapes suivantes devraient aider à résoudre tous les problèmes liés à **l’onglet Automate** n’apparaissant pas dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-146">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="a6b7d-147">[Assurez-vous que Microsoft 365 licence inclut Office scripts](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="a6b7d-147">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="a6b7d-148">[Vérifiez que votre navigateur est pris en charge](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="a6b7d-148">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="a6b7d-149">[Assurez-vous que les cookies tiers sont activés](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="a6b7d-149">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="a6b7d-150">[Assurez-vous que votre administrateur n’a pas désactivé Office scripts dans le Microsoft 365 d’administration](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="a6b7d-150">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="troubleshoot-scripts-in-power-automate"></a><span data-ttu-id="a6b7d-151">Scripts de dépannage dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="a6b7d-151">Troubleshoot scripts in Power Automate</span></span>

<span data-ttu-id="a6b7d-152">Pour plus d’informations spécifiques à l’exécution de scripts Power Automate, [consultez Les scripts de Office dépannage en cours d’exécution Power Automate](power-automate-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="a6b7d-152">For information specific to running scripts through Power Automate, see [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).</span></span>

## <a name="help-resources"></a><span data-ttu-id="a6b7d-153">Ressources d’aide</span><span class="sxs-lookup"><span data-stu-id="a6b7d-153">Help resources</span></span>

<span data-ttu-id="a6b7d-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) est une communauté de développeurs prêts à aider avec les problèmes de codage.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="a6b7d-155">Souvent, vous serez en mesure de trouver la solution à votre problème grâce à une recherche rapide Stack Overflow.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-155">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="a6b7d-156">Si ce n’est pas le cas, posez votre question et étiqueter avec l’étiquette « scripts de bureau ».</span><span class="sxs-lookup"><span data-stu-id="a6b7d-156">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="a6b7d-157">N’oubliez pas de mentionner que vous créez un script *Office,* pas un Office *Add-in*.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-157">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="a6b7d-158">Si vous rencontrez un problème avec l’API JavaScript Office, créez un problème dans le référentiel [officedev/office-js](https://github.com/OfficeDev/office-js) GitHub.s.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-158">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="a6b7d-159">Les membres de l’équipe produit répondront aux problèmes et fourniront une aide supplémentaire.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-159">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="a6b7d-160">La création d’un problème dans le référentiel **OfficeDev/office-js** indique que vous avez trouvé une faille dans la bibliothèque d’API JavaScript Office que l’équipe produit doit traiter.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-160">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="a6b7d-161">S’il y a un problème avec l’enregistreur d’action ou l’éditeur, envoyez des **commentaires via le bouton > d’aide** et de rétroaction Excel.</span><span class="sxs-lookup"><span data-stu-id="a6b7d-161">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="a6b7d-162">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a6b7d-162">See also</span></span>

- [<span data-ttu-id="a6b7d-163">Meilleures pratiques dans Office scripts</span><span class="sxs-lookup"><span data-stu-id="a6b7d-163">Best practices in Office Scripts</span></span>](../develop/best-practices.md)
- [<span data-ttu-id="a6b7d-164">Limites de plate-forme avec Office scripts</span><span class="sxs-lookup"><span data-stu-id="a6b7d-164">Platform limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="a6b7d-165">Améliorez les performances de vos scripts Office’argent</span><span class="sxs-lookup"><span data-stu-id="a6b7d-165">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="a6b7d-166">Scripts de Office en cours d’exécution dans PowerAutomate</span><span class="sxs-lookup"><span data-stu-id="a6b7d-166">Troubleshoot Office Scripts running in PowerAutomate</span></span>](power-automate-troubleshooting.md)
- [<span data-ttu-id="a6b7d-167">Annuler les effets des scripts Office texte</span><span class="sxs-lookup"><span data-stu-id="a6b7d-167">Undo the effects of Office Scripts</span></span>](undo.md)
