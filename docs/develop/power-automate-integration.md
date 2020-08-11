---
title: Exécuter des scripts Office avec Power automate
description: Comment obtenir des scripts Office pour Excel sur le Web avec un flux de travail Automated Power.
ms.date: 07/24/2020
localization_priority: Normal
ms.openlocfilehash: a427948847d7ab84962cdede7fb44d214592909f
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616674"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="f2421-103">Exécuter des scripts Office avec Power automate</span><span class="sxs-lookup"><span data-stu-id="f2421-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="f2421-104">[Power automate](https://flow.microsoft.com) vous permet d’ajouter des scripts Office à un flux de travail automatisé plus important.</span><span class="sxs-lookup"><span data-stu-id="f2421-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="f2421-105">Vous pouvez utiliser Power automate effectuer des opérations comme ajouter le contenu d’un message électronique à la table d’une feuille de calcul ou créer des actions dans vos outils de gestion de projet en fonction des commentaires de votre classeur.</span><span class="sxs-lookup"><span data-stu-id="f2421-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="f2421-106">Prise en main</span><span class="sxs-lookup"><span data-stu-id="f2421-106">Getting started</span></span>

<span data-ttu-id="f2421-107">Si vous ne connaissez pas l’automate de puissance, nous vous recommandons de consulter la [prise en main de Power automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="f2421-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="f2421-108">Ici, vous pouvez en savoir plus sur toutes les possibilités d’automatisation disponibles.</span><span class="sxs-lookup"><span data-stu-id="f2421-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="f2421-109">Les documents ici se concentrent sur la façon dont les scripts Office fonctionnent avec automate d’alimentation et sur la façon d’améliorer votre expérience Excel.</span><span class="sxs-lookup"><span data-stu-id="f2421-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="f2421-110">Pour commencer à combiner les scripts Power Automated et Office, suivez le didacticiel [commencer à utiliser des scripts avec Power automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="f2421-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="f2421-111">Cela vous apprend à créer un flux qui appelle un script simple.</span><span class="sxs-lookup"><span data-stu-id="f2421-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="f2421-112">Une fois que vous avez terminé ce didacticiel et que vous avez [transmis des données à des scripts dans un didacticiel de mise à niveau automatique de l’alimentation automatique](../tutorials/excel-power-automate-trigger.md) , renvoyez ici pour obtenir des informations détaillées sur la connexion de scripts Office à la mise à niveau automatique des flux.</span><span class="sxs-lookup"><span data-stu-id="f2421-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="f2421-113">Connecteur Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="f2421-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="f2421-114">Les [connecteurs](/connectors/connectors) sont les ponts entre l’automate de puissance et les applications.</span><span class="sxs-lookup"><span data-stu-id="f2421-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="f2421-115">Le [connecteur Excel Online (Business)](/connectors/excelonlinebusiness) donne accès à vos flux aux classeurs Excel.</span><span class="sxs-lookup"><span data-stu-id="f2421-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="f2421-116">L’action « exécuter un script » vous permet d’appeler n’importe quel script Office accessible via le classeur sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f2421-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="f2421-117">Vous pouvez également attribuer des paramètres d’entrée de scripts afin que les données puissent être fournies par le flux, ou que votre script renvoie des informations pour les étapes ultérieures dans le flux.</span><span class="sxs-lookup"><span data-stu-id="f2421-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f2421-118">L’action « exécuter un script » permet aux personnes qui utilisent le connecteur Excel d’accéder à votre classeur et à ses données.</span><span class="sxs-lookup"><span data-stu-id="f2421-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="f2421-119">De plus, il existe des risques de sécurité pour les scripts qui effectuent des appels d’API externes, comme expliqué dans la rubrique [appels externes de Power Automated](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="f2421-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="f2421-120">Si votre administrateur est concerné par l’exposition de données hautement sensibles, il peut soit désactiver le connecteur Excel Online, soit restreindre l’accès aux scripts Office via les contrôles de l' [administrateur des scripts Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="f2421-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="f2421-121">Transfert de données dans les flux pour les scripts</span><span class="sxs-lookup"><span data-stu-id="f2421-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="f2421-122">Power automate vous permet de transmettre des éléments de données entre les étapes de votre flux.</span><span class="sxs-lookup"><span data-stu-id="f2421-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="f2421-123">Les scripts peuvent être configurés pour accepter tous les types d’informations dont vous avez besoin et renvoyer tout élément de votre classeur souhaité dans votre flux.</span><span class="sxs-lookup"><span data-stu-id="f2421-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="f2421-124">L’entrée de votre script est spécifiée en ajoutant des paramètres à la `main` fonction (en plus de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="f2421-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="f2421-125">La sortie du script est déclarée en ajoutant un type de retour à `main` .</span><span class="sxs-lookup"><span data-stu-id="f2421-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="f2421-126">Lorsque vous créez un bloc de script d’exécution dans votre flux, les paramètres acceptés et les types renvoyés sont renseignés.</span><span class="sxs-lookup"><span data-stu-id="f2421-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="f2421-127">Si vous modifiez les paramètres ou les types de retour de votre script, vous devez rétablir le bloc de script « exécuter le script » de votre flux.</span><span class="sxs-lookup"><span data-stu-id="f2421-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="f2421-128">Cela permet de s’assurer que les données sont analysées correctement.</span><span class="sxs-lookup"><span data-stu-id="f2421-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="f2421-129">Les sections suivantes couvrent les détails de l’entrée et de la sortie des scripts utilisés dans Power automate.</span><span class="sxs-lookup"><span data-stu-id="f2421-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="f2421-130">Si vous souhaitez obtenir une approche pratique de l’apprentissage de cette rubrique, essayez de [transmettre des données à des scripts dans un didacticiel de puissance automate d’alimentation automatique](../tutorials/excel-power-automate-trigger.md) ou explorez le scénario d’exemple de [rappels de tâche automatisée](../resources/scenarios/task-reminders.md) .</span><span class="sxs-lookup"><span data-stu-id="f2421-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="f2421-131">`main`Paramètres : transmission de données à un script</span><span class="sxs-lookup"><span data-stu-id="f2421-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="f2421-132">Toutes les entrées de script sont spécifiées comme paramètres supplémentaires pour la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="f2421-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="f2421-133">Par exemple, si vous souhaitez qu’un script accepte un `string` qui représente un nom comme entrée, vous devez remplacer la `main` signature par `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="f2421-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="f2421-134">Lorsque vous configurez un flux dans Power Automated, vous pouvez spécifier des entrées de script sous forme de valeurs statiques, d' [expressions](/power-automate/use-expressions-in-conditions)ou de contenu dynamique.</span><span class="sxs-lookup"><span data-stu-id="f2421-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="f2421-135">Pour plus d’informations sur le connecteur d’un service individuel, consultez la [documentation relative](/connectors/)à la mise à niveau automatique du connecteur.</span><span class="sxs-lookup"><span data-stu-id="f2421-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="f2421-136">Lors de l’ajout de paramètres d’entrée à la fonction d’un script `main` , tenez compte des quotas et des restrictions suivantes.</span><span class="sxs-lookup"><span data-stu-id="f2421-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="f2421-137">Le premier paramètre doit être de type `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="f2421-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="f2421-138">Le nom de son paramètre n’a pas d’importance.</span><span class="sxs-lookup"><span data-stu-id="f2421-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="f2421-139">Chaque paramètre doit avoir un type (par exemple, `string` ou `number` ).</span><span class="sxs-lookup"><span data-stu-id="f2421-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="f2421-140">Les types de base,,,,, `string` `number` `boolean` `any` `unknown` `object` et `undefined` sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f2421-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="f2421-141">Les tableaux des types de base précédemment répertoriés sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f2421-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="f2421-142">Les tableaux imbriqués sont pris en charge en tant que paramètres (mais pas en tant que types de retour).</span><span class="sxs-lookup"><span data-stu-id="f2421-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="f2421-143">Les types Union sont autorisés s’ils sont une Union de littéraux appartenant à un même type (tel que `"Left" | "Right"` ).</span><span class="sxs-lookup"><span data-stu-id="f2421-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="f2421-144">Les unions d’un type pris en charge avec undefined sont également prises en charge (par exemple, `string | undefined` ).</span><span class="sxs-lookup"><span data-stu-id="f2421-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="f2421-145">Les types d’objet sont autorisés s’ils contiennent des propriétés de type `string` ,,, des `number` `boolean` tableaux pris en charge ou d’autres objets pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f2421-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="f2421-146">L’exemple suivant montre les objets imbriqués pris en charge en tant que types de paramètres :</span><span class="sxs-lookup"><span data-stu-id="f2421-146">The following example shows nested objects that are supported as parameter types:</span></span>

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. <span data-ttu-id="f2421-147">La définition de l’interface ou de la classe des objets doit être définie dans le script.</span><span class="sxs-lookup"><span data-stu-id="f2421-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="f2421-148">Un objet peut également être défini de manière anonyme, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="f2421-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="f2421-149">Les paramètres facultatifs sont autorisés et peuvent être dénotés comme tels à l’aide du modificateur facultatif `?` (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="f2421-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="f2421-150">Les valeurs de paramètre par défaut sont autorisées (par exemple `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="f2421-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="f2421-151">Renvoi de données à partir d’un script</span><span class="sxs-lookup"><span data-stu-id="f2421-151">Returning data from a script</span></span>

<span data-ttu-id="f2421-152">Les scripts peuvent renvoyer des données à partir du classeur afin d’être utilisées en tant que contenu dynamique dans un flux automatique de l’alimentation.</span><span class="sxs-lookup"><span data-stu-id="f2421-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="f2421-153">Comme avec les paramètres d’entrée, Power automate place certaines restrictions sur le type de retour.</span><span class="sxs-lookup"><span data-stu-id="f2421-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="f2421-154">Les types de base,,, `string` `number` `boolean` `void` et `undefined` sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f2421-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="f2421-155">Les types d’Union utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="f2421-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="f2421-156">Les types tableau sont autorisés s’ils sont de type `string` , `number` ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="f2421-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="f2421-157">Elles sont également autorisées si le type est un type de littéral Union pris en charge ou pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f2421-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="f2421-158">Les types d’objets utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="f2421-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="f2421-159">Le typage implicite est pris en charge, mais il doit suivre les mêmes règles qu’un type défini.</span><span class="sxs-lookup"><span data-stu-id="f2421-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="f2421-160">Éviter d’utiliser des références relatives</span><span class="sxs-lookup"><span data-stu-id="f2421-160">Avoid using relative references</span></span>

<span data-ttu-id="f2421-161">Power automate exécute votre script dans le classeur Excel choisi de votre part.</span><span class="sxs-lookup"><span data-stu-id="f2421-161">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="f2421-162">Le classeur peut être fermé lorsque cela se produit.</span><span class="sxs-lookup"><span data-stu-id="f2421-162">The workbook might be closed when this happens.</span></span> <span data-ttu-id="f2421-163">Toutes les API qui s’appuient sur l’état actuel de l’utilisateur, telles que `Workbook.getActiveWorksheet` , échouent lorsqu’elles sont exécutées via Power Automated.</span><span class="sxs-lookup"><span data-stu-id="f2421-163">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="f2421-164">Lors de la conception de vos scripts, veillez à utiliser des références absolues pour les feuilles de calcul et les plages.</span><span class="sxs-lookup"><span data-stu-id="f2421-164">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="f2421-165">Les méthodes suivantes génèrent une erreur et échouent lorsqu’elles sont appelées à partir d’un script dans un flux d’automate de puissance.</span><span class="sxs-lookup"><span data-stu-id="f2421-165">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="f2421-166">Class</span><span class="sxs-lookup"><span data-stu-id="f2421-166">Class</span></span> | <span data-ttu-id="f2421-167">Méthode</span><span class="sxs-lookup"><span data-stu-id="f2421-167">Method</span></span> |
|--|--|
| [<span data-ttu-id="f2421-168">Graphique</span><span class="sxs-lookup"><span data-stu-id="f2421-168">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="f2421-169">Range</span><span class="sxs-lookup"><span data-stu-id="f2421-169">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="f2421-170">Classeur</span><span class="sxs-lookup"><span data-stu-id="f2421-170">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="f2421-171">Classeur</span><span class="sxs-lookup"><span data-stu-id="f2421-171">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="f2421-172">Classeur</span><span class="sxs-lookup"><span data-stu-id="f2421-172">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="f2421-173">Classeur</span><span class="sxs-lookup"><span data-stu-id="f2421-173">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [<span data-ttu-id="f2421-174">Classeur</span><span class="sxs-lookup"><span data-stu-id="f2421-174">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="f2421-175">Classeur</span><span class="sxs-lookup"><span data-stu-id="f2421-175">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [<span data-ttu-id="f2421-176">Feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="f2421-176">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a><span data-ttu-id="f2421-177">Exemple</span><span class="sxs-lookup"><span data-stu-id="f2421-177">Example</span></span>

<span data-ttu-id="f2421-178">La capture d’écran suivante montre un flux automatique de puissance déclenché à chaque fois qu’un problème [GitHub](https://github.com/) vous est affecté.</span><span class="sxs-lookup"><span data-stu-id="f2421-178">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="f2421-179">Le flux exécute un script qui ajoute le problème à un tableau dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="f2421-179">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="f2421-180">Si ce tableau comporte au moins cinq problèmes, le flux envoie un rappel par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="f2421-180">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![Exemple de flux tel qu’illustré dans l’éditeur de flux Automated Power.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="f2421-182">La `main` fonction du script spécifie l’ID du problème et le titre du problème en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans la table des problèmes.</span><span class="sxs-lookup"><span data-stu-id="f2421-182">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a><span data-ttu-id="f2421-183">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f2421-183">See also</span></span>

- [<span data-ttu-id="f2421-184">Exécuter des scripts Office dans Excel sur le Web avec Power Automated Power</span><span class="sxs-lookup"><span data-stu-id="f2421-184">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="f2421-185">Transmettre des données à des scripts dans un flux automatique Power Automate</span><span class="sxs-lookup"><span data-stu-id="f2421-185">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="f2421-186">Principes de base pour la rédaction de scripts Office en Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="f2421-186">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="f2421-187">Prise en main de Power Automate</span><span class="sxs-lookup"><span data-stu-id="f2421-187">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="f2421-188">Documentation de référence du connecteur Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="f2421-188">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
