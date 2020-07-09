---
title: Exécuter des scripts Office avec Power automate
description: Comment obtenir des scripts Office pour Excel sur le Web avec un flux de travail Automated Power.
ms.date: 07/01/2020
localization_priority: Normal
ms.openlocfilehash: 40a67f3d0e8f049a8ec5516c0af54c5fc6fb9319
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081592"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="10098-103">Exécuter des scripts Office avec Power automate</span><span class="sxs-lookup"><span data-stu-id="10098-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="10098-104">[Power automate](https://flow.microsoft.com) vous permet d’ajouter des scripts Office à un flux de travail automatisé plus important.</span><span class="sxs-lookup"><span data-stu-id="10098-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="10098-105">Vous pouvez utiliser Power automate effectuer des opérations comme ajouter le contenu d’un message électronique à la table d’une feuille de calcul ou créer des actions dans vos outils de gestion de projet en fonction des commentaires de votre classeur.</span><span class="sxs-lookup"><span data-stu-id="10098-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="10098-106">Si vous ne connaissez pas l’automate de puissance, nous vous recommandons de consulter la [prise en main de Power automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="10098-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="10098-107">Ici, vous pouvez en savoir plus sur l’automatisation de vos flux de travail sur plusieurs services.</span><span class="sxs-lookup"><span data-stu-id="10098-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="10098-108">Actuellement, vous ne pouvez pas exécuter des scripts Office à partir d’un [flux partagé](/power-automate/share-buttons).</span><span class="sxs-lookup"><span data-stu-id="10098-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="10098-109">Seul l’utilisateur qui a créé un script peut l’exécuter, même via automate d’alimentation.</span><span class="sxs-lookup"><span data-stu-id="10098-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="10098-110">Prise en main</span><span class="sxs-lookup"><span data-stu-id="10098-110">Getting started</span></span>

<span data-ttu-id="10098-111">Pour commencer à combiner les scripts Power Automated et Office, suivez le didacticiel [commencer à utiliser des scripts avec Power automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="10098-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="10098-112">Cela vous apprend à créer un flux qui appelle un script simple.</span><span class="sxs-lookup"><span data-stu-id="10098-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="10098-113">Une fois que vous avez terminé ce didacticiel et que vous avez [exécuté automatiquement des scripts avec Automated Power](../tutorials/excel-power-automate-trigger.md) Automated Power Tutorial Tutorial, renvoyez ici pour obtenir des informations détaillées sur la connexion de scripts Office à Power Automated flows.</span><span class="sxs-lookup"><span data-stu-id="10098-113">After you've completed that tutorial and the [Automatically run scripts with automated Power Automate flows](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="10098-114">Connecteur Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="10098-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="10098-115">Les [connecteurs](/connectors/connectors) sont les ponts entre l’automate de puissance et les applications.</span><span class="sxs-lookup"><span data-stu-id="10098-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="10098-116">Le [connecteur Excel Online (Business)](/connectors/excelonlinebusiness) donne accès à vos flux aux classeurs Excel.</span><span class="sxs-lookup"><span data-stu-id="10098-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="10098-117">L’action « exécuter un script » vous permet d’appeler n’importe quel script Office accessible via le classeur sélectionné.</span><span class="sxs-lookup"><span data-stu-id="10098-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="10098-118">Vous pouvez non seulement exécuter des scripts via un flux, mais vous pouvez transmettre des données vers et depuis le classeur avec le flux via les scripts.</span><span class="sxs-lookup"><span data-stu-id="10098-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="10098-119">L’action « exécuter un script » permet aux personnes qui utilisent le connecteur Excel d’accéder à votre classeur et à ses données.</span><span class="sxs-lookup"><span data-stu-id="10098-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="10098-120">De plus, il existe des risques de sécurité pour les scripts qui effectuent des appels d’API externes, comme expliqué dans la rubrique [appels externes de Power Automated](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="10098-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="10098-121">Si votre administrateur est concerné par l’exposition de données hautement sensibles, il peut soit désactiver le connecteur Excel Online, soit restreindre l’accès aux scripts Office via les contrôles de l' [administrateur des scripts Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="10098-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="10098-122">Transfert de données dans les flux pour les scripts</span><span class="sxs-lookup"><span data-stu-id="10098-122">Data transfer in flows for scripts</span></span>

<span data-ttu-id="10098-123">Power automate vous permet de transmettre des éléments de données entre les étapes de votre flux.</span><span class="sxs-lookup"><span data-stu-id="10098-123">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="10098-124">Les scripts peuvent être configurés pour accepter tous les types d’informations dont vous avez besoin et renvoyer tout élément de votre classeur souhaité dans votre flux.</span><span class="sxs-lookup"><span data-stu-id="10098-124">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="10098-125">L’entrée de votre script est spécifiée en ajoutant des paramètres à la `main` fonction (en plus de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="10098-125">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="10098-126">La sortie du script est déclarée en ajoutant un type de retour à `main` .</span><span class="sxs-lookup"><span data-stu-id="10098-126">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="10098-127">Lorsque vous créez un bloc de script d’exécution dans votre flux, les paramètres acceptés et les types renvoyés sont renseignés.</span><span class="sxs-lookup"><span data-stu-id="10098-127">When you create a "Run Script" block in you flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="10098-128">Si vous modifiez les paramètres ou les types de retour de votre script, vous devez rétablir le bloc de script « exécuter le script » de votre flux.</span><span class="sxs-lookup"><span data-stu-id="10098-128">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="10098-129">Cela garantit que les données sont analysées correctement.</span><span class="sxs-lookup"><span data-stu-id="10098-129">This ensure the data is being parsed correctly.</span></span>

<span data-ttu-id="10098-130">Les sections suivantes couvrent les détails de l’entrée et de la sortie des scripts utilisés dans Power automate.</span><span class="sxs-lookup"><span data-stu-id="10098-130">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="10098-131">Si vous souhaitez une approche pratique de l’apprentissage de cette rubrique, essayez les [scripts exécuter automatiquement avec Automated Power](../tutorials/excel-power-automate-trigger.md) Automated Flow Tutorial ou explorez le scénario d’exemple de [rappel de tâche automatisée](../resources/scenarios/task-reminders.md) .</span><span class="sxs-lookup"><span data-stu-id="10098-131">If you'd like a hands-on approach to learning this topic, try out the [Automatically run scripts with automated Power Automate flows](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="10098-132">`main`Paramètres : transmission de données à un script</span><span class="sxs-lookup"><span data-stu-id="10098-132">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="10098-133">Toutes les entrées de script sont spécifiées comme paramètres supplémentaires pour la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="10098-133">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="10098-134">Par exemple, si vous souhaitez qu’un script accepte un `string` qui représente un nom comme entrée, vous devez remplacer la `main` signature par `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="10098-134">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="10098-135">Lorsque vous configurez un flux dans Power Automated, vous pouvez spécifier des entrées de script sous forme de valeurs statiques, d' [expressions](/power-automate/use-expressions-in-conditions)ou de contenu dynamique.</span><span class="sxs-lookup"><span data-stu-id="10098-135">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="10098-136">Pour plus d’informations sur le connecteur d’un service individuel, consultez la [documentation relative](/connectors/)à la mise à niveau automatique du connecteur.</span><span class="sxs-lookup"><span data-stu-id="10098-136">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="10098-137">Lors de l’ajout de paramètres d’entrée à la fonction d’un script `main` , tenez compte des quotas et des restrictions suivantes.</span><span class="sxs-lookup"><span data-stu-id="10098-137">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="10098-138">Le premier paramètre doit être de type `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="10098-138">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="10098-139">Le nom de son paramètre n’a pas d’importance.</span><span class="sxs-lookup"><span data-stu-id="10098-139">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="10098-140">Chaque paramètre doit avoir un type.</span><span class="sxs-lookup"><span data-stu-id="10098-140">Every parameter must have a type.</span></span>

3. <span data-ttu-id="10098-141">Les types de base,,,,, `string` `number` `boolean` `any` `unknown` `object` et `undefined` sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="10098-141">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="10098-142">Les tableaux des types de base précédemment répertoriés sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="10098-142">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="10098-143">Les tableaux imbriqués sont pris en charge en tant que paramètres (mais pas en tant que types de retour).</span><span class="sxs-lookup"><span data-stu-id="10098-143">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="10098-144">Les types Union sont autorisés s’il s’agit d’une Union de littéraux appartenant à un seul type ( `string` , `number` , ou `boolean` ).</span><span class="sxs-lookup"><span data-stu-id="10098-144">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="10098-145">Les unions d’un type pris en charge avec undefined sont également prises en charge.</span><span class="sxs-lookup"><span data-stu-id="10098-145">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="10098-146">Les types d’objet sont autorisés s’ils contiennent des propriétés de type `string` ,,, des `number` `boolean` tableaux pris en charge ou d’autres objets pris en charge.</span><span class="sxs-lookup"><span data-stu-id="10098-146">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="10098-147">L’exemple suivant montre les objets imbriqués pris en charge en tant que types de paramètres :</span><span class="sxs-lookup"><span data-stu-id="10098-147">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="10098-148">La définition de l’interface ou de la classe des objets doit être définie dans le script.</span><span class="sxs-lookup"><span data-stu-id="10098-148">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="10098-149">Un objet peut également être défini de manière anonyme, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="10098-149">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="10098-150">Les paramètres facultatifs sont autorisés et peuvent être dénotés comme tels à l’aide du modificateur facultatif `?` (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="10098-150">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="10098-151">Les valeurs de paramètre par défaut sont autorisées (par exemple `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="10098-151">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script"></a><span data-ttu-id="10098-152">Renvoi de données à partir d’un script</span><span class="sxs-lookup"><span data-stu-id="10098-152">Returning data from a script</span></span>

<span data-ttu-id="10098-153">Les scripts peuvent renvoyer des données à partir du classeur afin d’être utilisées en tant que contenu dynamique dans un flux automatique de l’alimentation.</span><span class="sxs-lookup"><span data-stu-id="10098-153">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="10098-154">Comme avec les paramètres d’entrée, Power automate place certaines restrictions sur le type de retour.</span><span class="sxs-lookup"><span data-stu-id="10098-154">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="10098-155">Les types de base,,, `string` `number` `boolean` `void` et `undefined` sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="10098-155">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="10098-156">Les types d’Union utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="10098-156">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="10098-157">Les types tableau sont autorisés s’ils sont de type `string` , `number` ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="10098-157">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="10098-158">Elles sont également autorisées si le type est un type de littéral Union pris en charge ou pris en charge.</span><span class="sxs-lookup"><span data-stu-id="10098-158">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="10098-159">Les types d’objets utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="10098-159">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="10098-160">Le typage implicite est pris en charge, mais il doit suivre les mêmes règles qu’un type défini.</span><span class="sxs-lookup"><span data-stu-id="10098-160">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="10098-161">Éviter d’utiliser des références relatives</span><span class="sxs-lookup"><span data-stu-id="10098-161">Avoid using relative references</span></span>

<span data-ttu-id="10098-162">Power automate exécute votre script dans le classeur Excel choisi de votre part.</span><span class="sxs-lookup"><span data-stu-id="10098-162">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="10098-163">Le classeur peut être fermé lorsque cela se produit.</span><span class="sxs-lookup"><span data-stu-id="10098-163">The workbook might be closed when this happens.</span></span> <span data-ttu-id="10098-164">Toutes les API qui s’appuient sur l’état actuel de l’utilisateur, telles que `Workbook.getActiveWorksheet` , échouent lorsqu’elles sont exécutées via Power Automated.</span><span class="sxs-lookup"><span data-stu-id="10098-164">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="10098-165">Lors de la conception de vos scripts, veillez à utiliser des références absolues pour les feuilles de calcul et les plages.</span><span class="sxs-lookup"><span data-stu-id="10098-165">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="10098-166">Les fonctions suivantes génèrent une erreur et échouent lorsqu’elles sont appelées à partir d’un script dans un flux d’automate de puissance.</span><span class="sxs-lookup"><span data-stu-id="10098-166">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a><span data-ttu-id="10098-167">Exemple</span><span class="sxs-lookup"><span data-stu-id="10098-167">Example</span></span>

<span data-ttu-id="10098-168">La capture d’écran suivante montre un flux automatique de puissance déclenché à chaque fois qu’un problème [GitHub](https://github.com/) vous est affecté.</span><span class="sxs-lookup"><span data-stu-id="10098-168">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="10098-169">Le flux exécute un script qui ajoute le problème à un tableau dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="10098-169">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="10098-170">Si ce tableau comporte au moins cinq problèmes, le flux envoie un rappel par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="10098-170">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![Exemple de flux tel qu’illustré dans l’éditeur de flux Automated Power.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="10098-172">La `main` fonction du script spécifie l’ID du problème et le titre du problème en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans la table des problèmes.</span><span class="sxs-lookup"><span data-stu-id="10098-172">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="10098-173">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="10098-173">See also</span></span>

- [<span data-ttu-id="10098-174">Exécuter des scripts Office dans Excel sur le Web avec Power Automated Power</span><span class="sxs-lookup"><span data-stu-id="10098-174">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="10098-175">Exécuter automatiquement des scripts avec automate d’alimentation automatisée des flux</span><span class="sxs-lookup"><span data-stu-id="10098-175">Automatically run scripts with automated Power Automate flows</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="10098-176">Principes de base des scripts pour Office Scripts dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="10098-176">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="10098-177">Prise en main de Power Automate</span><span class="sxs-lookup"><span data-stu-id="10098-177">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="10098-178">Documentation de référence du connecteur Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="10098-178">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
