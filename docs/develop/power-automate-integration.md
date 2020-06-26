---
title: Intégration de scripts Office avec Power Automated
description: Comment obtenir des scripts Office pour Excel sur le Web avec un flux de travail Automated Power.
ms.date: 06/24/2020
localization_priority: Normal
ms.openlocfilehash: 977d9c88d75c8070eb729a443b4e8bc9a32e456d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878775"
---
# <a name="integrate-office-scripts-with-power-automate"></a><span data-ttu-id="cb1b5-103">Intégration de scripts Office avec Power Automated</span><span class="sxs-lookup"><span data-stu-id="cb1b5-103">Integrate Office Scripts with Power Automate</span></span>

<span data-ttu-id="cb1b5-104">[Power automate](https://flow.microsoft.com) intègre votre script dans un flux de travail plus important.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-104">[Power Automate](https://flow.microsoft.com) integrates your script into a larger workflow.</span></span> <span data-ttu-id="cb1b5-105">Vous pouvez utiliser Power automate effectuer des opérations comme ajouter le contenu d’un message électronique à la table d’une feuille de calcul ou créer des actions dans vos outils de gestion de projet en fonction des commentaires de votre classeur.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="cb1b5-106">Si vous ne connaissez pas l’automate de puissance, nous vous recommandons de consulter la [prise en main de Power automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="cb1b5-107">Ici, vous pouvez en savoir plus sur l’automatisation de vos flux de travail sur plusieurs services.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cb1b5-108">Actuellement, vous ne pouvez pas exécuter des scripts Office à partir d’un [flux partagé](/power-automate/share-buttons).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="cb1b5-109">Seul l’utilisateur qui a créé un script peut l’exécuter, même via automate d’alimentation.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="cb1b5-110">Prise en main</span><span class="sxs-lookup"><span data-stu-id="cb1b5-110">Getting started</span></span>

<span data-ttu-id="cb1b5-111">Pour commencer à combiner les scripts Power Automated et Office, suivez le didacticiel [commencer à utiliser des scripts avec Power automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="cb1b5-112">Cela vous apprend à créer un flux qui appelle un script simple.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="cb1b5-113">Une fois que vous avez terminé ce didacticiel et que vous avez [exécuté automatiquement les scripts avec Power Automated](../tutorials/excel-power-automate-trigger.md) , revenez ici pour en savoir plus sur les intégrations de plateforme.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-113">After you've completed that tutorial and the [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md) tutorial, return here to learn details about the platform integrations.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="cb1b5-114">Connecteur Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="cb1b5-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="cb1b5-115">Les [connecteurs](/connectors/connectors) sont les ponts entre l’automate de puissance et les applications.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="cb1b5-116">Le [connecteur Excel Online (Business)](/connectors/excelonlinebusiness) donne accès à vos flux aux classeurs Excel.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="cb1b5-117">L’action « exécuter un script » vous permet d’appeler n’importe quel script Office accessible via le classeur sélectionné.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="cb1b5-118">Vous pouvez non seulement exécuter des scripts via un flux, mais vous pouvez transmettre des données vers et depuis le classeur avec le flux via les scripts.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cb1b5-119">L’action « exécuter un script » permet aux personnes qui utilisent le connecteur Excel d’accéder à votre classeur et à ses données.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="cb1b5-120">De plus, il existe des risques de sécurité pour les scripts qui effectuent des appels d’API externes, comme expliqué dans la rubrique [appels externes de Power Automated](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="cb1b5-121">Si votre administrateur est concerné par l’exposition de données hautement sensibles, il peut soit désactiver le connecteur Excel Online, soit restreindre l’accès aux scripts Office via les contrôles de l' [administrateur des scripts Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="passing-data-from-power-automate-into-a-script"></a><span data-ttu-id="cb1b5-122">Transmission de données de Power automate dans un script</span><span class="sxs-lookup"><span data-stu-id="cb1b5-122">Passing data from Power Automate into a script</span></span>

<span data-ttu-id="cb1b5-123">Toutes les entrées de script sont spécifiées comme paramètres supplémentaires pour la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-123">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="cb1b5-124">Par exemple, si vous souhaitez qu’un script accepte un `string` qui représente un nom comme entrée, vous devez remplacer la `main` signature par `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="cb1b5-124">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="cb1b5-125">Lorsque vous configurez un flux dans Power Automated, vous pouvez spécifier des entrées de script sous forme de valeurs statiques, d' [expressions](/power-automate/use-expressions-in-conditions)ou de contenu dynamique.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-125">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="cb1b5-126">Pour plus d’informations sur le connecteur d’un service individuel, consultez la [documentation relative](/connectors/)à la mise à niveau automatique du connecteur.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-126">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="cb1b5-127">Lors de l’ajout de paramètres d’entrée à la fonction d’un script `main` , tenez compte des quotas et des restrictions suivantes.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-127">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="cb1b5-128">Le premier paramètre doit être de type `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="cb1b5-128">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="cb1b5-129">Le nom de son paramètre n’a pas d’importance.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-129">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="cb1b5-130">Chaque paramètre doit avoir un type.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-130">Every parameter must have a type.</span></span>

3. <span data-ttu-id="cb1b5-131">Les types de base,,,,, `string` `number` `boolean` `any` `unknown` `object` et `undefined` sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-131">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="cb1b5-132">Les tableaux des types de base précédemment répertoriés sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-132">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="cb1b5-133">Les tableaux imbriqués sont pris en charge en tant que paramètres (mais pas en tant que types de retour).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-133">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="cb1b5-134">Les types Union sont autorisés s’il s’agit d’une Union de littéraux appartenant à un seul type ( `string` , `number` , ou `boolean` ).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-134">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="cb1b5-135">Les unions d’un type pris en charge avec undefined sont également prises en charge.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-135">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="cb1b5-136">Les types d’objet sont autorisés s’ils contiennent des propriétés de type `string` ,,, des `number` `boolean` tableaux pris en charge ou d’autres objets pris en charge.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-136">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="cb1b5-137">L’exemple suivant montre les objets imbriqués pris en charge en tant que types de paramètres :</span><span class="sxs-lookup"><span data-stu-id="cb1b5-137">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="cb1b5-138">La définition de l’interface ou de la classe des objets doit être définie dans le script.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-138">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="cb1b5-139">Un objet peut également être défini de manière anonyme, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="cb1b5-139">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="cb1b5-140">Les paramètres facultatifs sont autorisés et peuvent être dénotés comme tels à l’aide du modificateur facultatif `?` (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="cb1b5-140">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="cb1b5-141">Les valeurs de paramètre par défaut sont autorisées (par exemple `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="cb1b5-141">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script-back-to-power-automate"></a><span data-ttu-id="cb1b5-142">Retour des données d’un script à automate d’alimentation</span><span class="sxs-lookup"><span data-stu-id="cb1b5-142">Returning data from a script back to Power Automate</span></span>

<span data-ttu-id="cb1b5-143">Les scripts peuvent renvoyer des données à partir du classeur afin d’être utilisées en tant que contenu dynamique dans un flux automatique de l’alimentation.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-143">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="cb1b5-144">Comme avec les paramètres d’entrée, Power automate place certaines restrictions sur le type de retour.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-144">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="cb1b5-145">Les types de base,,, `string` `number` `boolean` `void` et `undefined` sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-145">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="cb1b5-146">Les types d’Union utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-146">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="cb1b5-147">Les types tableau sont autorisés s’ils sont de type `string` , `number` ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="cb1b5-147">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="cb1b5-148">Elles sont également autorisées si le type est un type de littéral Union pris en charge ou pris en charge.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-148">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="cb1b5-149">Les types d’objets utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-149">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="cb1b5-150">Le typage implicite est pris en charge, mais il doit suivre les mêmes règles qu’un type défini.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-150">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="cb1b5-151">Éviter d’utiliser des références relatives</span><span class="sxs-lookup"><span data-stu-id="cb1b5-151">Avoid using relative references</span></span>

<span data-ttu-id="cb1b5-152">Power automate exécute votre script dans le classeur Excel choisi de votre part.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-152">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="cb1b5-153">Le classeur peut être fermé lorsque cela se produit.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-153">The workbook might be closed when this happens.</span></span> <span data-ttu-id="cb1b5-154">Toutes les API qui s’appuient sur l’état actuel de l’utilisateur, telles que `Workbook.getActiveWorksheet` , échouent lorsqu’elles sont exécutées via Power Automated.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-154">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="cb1b5-155">Lors de la conception de vos scripts, veillez à utiliser des références absolues pour les feuilles de calcul et les plages.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-155">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="cb1b5-156">Les fonctions suivantes génèrent une erreur et échouent lorsqu’elles sont appelées à partir d’un script dans un flux d’automate de puissance.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-156">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

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

## <a name="example"></a><span data-ttu-id="cb1b5-157">Exemple</span><span class="sxs-lookup"><span data-stu-id="cb1b5-157">Example</span></span>

<span data-ttu-id="cb1b5-158">La capture d’écran suivante montre un flux automatique de puissance déclenché à chaque fois qu’un problème [GitHub](https://github.com/) vous est affecté.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-158">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="cb1b5-159">Le flux exécute un script qui ajoute le problème à un tableau dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-159">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="cb1b5-160">Si ce tableau comporte au moins cinq problèmes, le flux envoie un rappel par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-160">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![Exemple de flux tel qu’illustré dans l’éditeur de flux Automated Power.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="cb1b5-162">La `main` fonction du script spécifie l’ID du problème et le titre du problème en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans la table des problèmes.</span><span class="sxs-lookup"><span data-stu-id="cb1b5-162">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="cb1b5-163">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cb1b5-163">See also</span></span>

- [<span data-ttu-id="cb1b5-164">Exécuter des scripts Office dans Excel sur le Web avec Power Automated Power</span><span class="sxs-lookup"><span data-stu-id="cb1b5-164">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="cb1b5-165">Exécuter automatiquement des scripts avec Power automate</span><span class="sxs-lookup"><span data-stu-id="cb1b5-165">Automatically run scripts with Power Automate</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="cb1b5-166">Principes de base des scripts pour Office Scripts dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="cb1b5-166">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="cb1b5-167">Prise en main de Power Automate</span><span class="sxs-lookup"><span data-stu-id="cb1b5-167">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="cb1b5-168">Documentation de référence du connecteur Excel Online (Business)</span><span class="sxs-lookup"><span data-stu-id="cb1b5-168">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
