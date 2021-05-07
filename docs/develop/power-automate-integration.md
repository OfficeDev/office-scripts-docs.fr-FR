---
title: Exécuter Office scripts avec Power Automate
description: Comment obtenir des scripts Office pour Excel sur le Web un flux de travail Power Automate de travail.
ms.date: 12/16/2020
localization_priority: Normal
ms.openlocfilehash: fd2622880f08c253f4333e642d1ebb0410bce681
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232416"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="fa03d-103">Exécuter Office scripts avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="fa03d-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="fa03d-104">[Power Automate](https://flow.microsoft.com) vous permet d’ajouter Office scripts à un flux de travail automatisé plus important.</span><span class="sxs-lookup"><span data-stu-id="fa03d-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="fa03d-105">Vous pouvez utiliser Power Automate des opérations telles que l’ajout du contenu d’un e-mail au tableau d’une feuille de calcul ou la création d’actions dans vos outils de gestion de projet en fonction des commentaires de votre feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="fa03d-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="getting-started"></a><span data-ttu-id="fa03d-106">Prise en main</span><span class="sxs-lookup"><span data-stu-id="fa03d-106">Getting started</span></span>

<span data-ttu-id="fa03d-107">Si vous débutez avec Power Automate, nous vous recommandons de visiter La mise en [Power Automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="fa03d-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="fa03d-108">Vous y découvrirez toutes les possibilités d’automatisation disponibles.</span><span class="sxs-lookup"><span data-stu-id="fa03d-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="fa03d-109">Les documents présentés ici se concentrent sur Office fonctionnement des scripts Power Automate et sur la façon dont cela peut vous aider à améliorer Excel expérience utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fa03d-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="fa03d-110">Pour commencer à combiner Power Automate et Office scripts, suivez le didacticiel Démarrer à l’aide de [scripts Power Automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="fa03d-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="fa03d-111">Cela vous montre comment créer un flux qui appelle un script simple.</span><span class="sxs-lookup"><span data-stu-id="fa03d-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="fa03d-112">Une fois que vous avez terminé ce didacticiel et passé les données aux [scripts](../tutorials/excel-power-automate-trigger.md) dans un didacticiel de flux Power Automate exécuté automatiquement, revenir ici pour obtenir des informations détaillées sur la connexion de scripts Office à des flux Power Automate.</span><span class="sxs-lookup"><span data-stu-id="fa03d-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="fa03d-113">Excel Connecteur En ligne (Entreprise)</span><span class="sxs-lookup"><span data-stu-id="fa03d-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="fa03d-114">[Les connecteurs](/connectors/connectors) sont les ponts entre Power Automate applications.</span><span class="sxs-lookup"><span data-stu-id="fa03d-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="fa03d-115">Le [connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness) permet à vos flux d’accéder Excel de travail.</span><span class="sxs-lookup"><span data-stu-id="fa03d-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="fa03d-116">L’action « Exécuter le script » vous permet d’appeler Office script accessible via le livre de travail sélectionné.</span><span class="sxs-lookup"><span data-stu-id="fa03d-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="fa03d-117">Vous pouvez également donner à vos scripts des paramètres d’entrée afin que les données soient fournies par le flux ou que votre script retourne des informations pour les étapes ultérieures du flux.</span><span class="sxs-lookup"><span data-stu-id="fa03d-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fa03d-118">L’action « Exécuter le script » donne aux utilisateurs du connecteur Excel un accès significatif à votre workbook et à ses données.</span><span class="sxs-lookup"><span data-stu-id="fa03d-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="fa03d-119">En outre, il existe des risques de sécurité avec les scripts qui font des appels d’API externes, comme expliqué dans les appels externes de [Power Automate](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="fa03d-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="fa03d-120">Si votre administrateur est préoccupé par l’exposition de données hautement sensibles, il peut désactiver le connecteur Excel Online ou restreindre l’accès aux scripts Office par le biais des contrôles d’administrateur [Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="fa03d-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="fa03d-121">Transfert de données dans les flux pour les scripts</span><span class="sxs-lookup"><span data-stu-id="fa03d-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="fa03d-122">Power Automate vous permet de passer des éléments de données entre les étapes de votre flux.</span><span class="sxs-lookup"><span data-stu-id="fa03d-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="fa03d-123">Les scripts peuvent être configurés pour accepter les types d’informations dont vous avez besoin et renvoyer tout ce dont vous avez besoin dans votre flux de travail.</span><span class="sxs-lookup"><span data-stu-id="fa03d-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="fa03d-124">L’entrée de votre script est spécifiée en ajoutant des paramètres à la `main` fonction (en plus de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="fa03d-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="fa03d-125">La sortie du script est déclarée en ajoutant un type de retour à `main` .</span><span class="sxs-lookup"><span data-stu-id="fa03d-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="fa03d-126">Lorsque vous créez un bloc « Exécuter un script » dans votre flux, les paramètres acceptés et les types renvoyés sont remplis.</span><span class="sxs-lookup"><span data-stu-id="fa03d-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="fa03d-127">Si vous modifiez les paramètres ou renvoyez des types de votre script, vous devrez revenir au bloc « Exécuter le script » de votre flux.</span><span class="sxs-lookup"><span data-stu-id="fa03d-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="fa03d-128">Cela garantit que les données sont en cours d’analyse correctement.</span><span class="sxs-lookup"><span data-stu-id="fa03d-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="fa03d-129">Les sections suivantes couvrent les détails de l’entrée et de la sortie pour les scripts utilisés dans Power Automate.</span><span class="sxs-lookup"><span data-stu-id="fa03d-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="fa03d-130">Si vous souhaitez une approche pratique de l’apprentissage de cette rubrique, essayez de transmettre des données aux [scripts](../tutorials/excel-power-automate-trigger.md) dans un didacticiel de flux Power Automate exécuté automatiquement ou explorez l’exemple de scénario de [rappels](../resources/scenarios/task-reminders.md) de tâches automatisés.</span><span class="sxs-lookup"><span data-stu-id="fa03d-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-passing-data-to-a-script"></a><span data-ttu-id="fa03d-131">`main` Paramètres : transmission de données à un script</span><span class="sxs-lookup"><span data-stu-id="fa03d-131">`main` Parameters: Passing data to a script</span></span>

<span data-ttu-id="fa03d-132">Toutes les entrées de script sont spécifiées en tant que paramètres supplémentaires pour la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="fa03d-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="fa03d-133">Par exemple, si vous souhaitez qu’un script accepte un nom qui représente un nom comme entrée, vous devez modifier `string` la `main` signature en `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="fa03d-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="fa03d-134">Lorsque vous configurez un flux dans Power Automate, vous pouvez spécifier une entrée de script en tant que valeurs statiques, [expressions](/power-automate/use-expressions-in-conditions)ou contenu dynamique.</span><span class="sxs-lookup"><span data-stu-id="fa03d-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="fa03d-135">Pour plus d’informations sur le connecteur d’un service individuel, voir la [documentation Power Automate Connector.](/connectors/)</span><span class="sxs-lookup"><span data-stu-id="fa03d-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="fa03d-136">Lorsque vous ajoutez des paramètres d’entrée à la fonction d’un script, prenons en compte les `main` limites et restrictions suivantes.</span><span class="sxs-lookup"><span data-stu-id="fa03d-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="fa03d-137">Le premier paramètre doit être de type `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="fa03d-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="fa03d-138">Son nom de paramètre n’a pas d’importance.</span><span class="sxs-lookup"><span data-stu-id="fa03d-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="fa03d-139">Chaque paramètre doit avoir un type (par `string` exemple, ou `number` ).</span><span class="sxs-lookup"><span data-stu-id="fa03d-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="fa03d-140">Les types `string` de base , , , , et sont pris en `number` `boolean` `any` `unknown` `object` `undefined` charge.</span><span class="sxs-lookup"><span data-stu-id="fa03d-140">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="fa03d-141">Les tableaux des types de base répertoriés précédemment sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="fa03d-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="fa03d-142">Les tableaux imbrmbrés sont pris en charge en tant que paramètres (mais pas en tant que types de retour).</span><span class="sxs-lookup"><span data-stu-id="fa03d-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="fa03d-143">Les types Union sont autorisés s’il s’agit d’une union de littéraux appartenant à un seul type (par `"Left" | "Right"` exemple).</span><span class="sxs-lookup"><span data-stu-id="fa03d-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="fa03d-144">Les personnes d’un type pris en charge avec undefined sont également pris en charge (par `string | undefined` exemple).</span><span class="sxs-lookup"><span data-stu-id="fa03d-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="fa03d-145">Les types d’objets sont autorisés s’ils contiennent des propriétés de type , , tableaux pris `string` en charge ou autres objets pris en `number` `boolean` charge.</span><span class="sxs-lookup"><span data-stu-id="fa03d-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="fa03d-146">L’exemple suivant montre les objets imbrmbrés pris en charge en tant que types de paramètres :</span><span class="sxs-lookup"><span data-stu-id="fa03d-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="fa03d-147">L’interface ou la définition de classe des objets doit être définie dans le script.</span><span class="sxs-lookup"><span data-stu-id="fa03d-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="fa03d-148">Un objet peut également être défini de manière anonyme en ligne, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="fa03d-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="fa03d-149">Les paramètres facultatifs sont autorisés et peuvent être indiqués en tant que tels à l’aide du modificateur facultatif `?` (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="fa03d-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="fa03d-150">Les valeurs de paramètre par défaut sont autorisées (par `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` exemple.</span><span class="sxs-lookup"><span data-stu-id="fa03d-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="returning-data-from-a-script"></a><span data-ttu-id="fa03d-151">Renvoi de données à partir d’un script</span><span class="sxs-lookup"><span data-stu-id="fa03d-151">Returning data from a script</span></span>

<span data-ttu-id="fa03d-152">Les scripts peuvent renvoyer des données à partir du workbook à utiliser en tant que contenu dynamique dans Power Automate flux.</span><span class="sxs-lookup"><span data-stu-id="fa03d-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="fa03d-153">Comme pour les paramètres d’entrée, Power Automate des restrictions sur le type de retour.</span><span class="sxs-lookup"><span data-stu-id="fa03d-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="fa03d-154">Les types `string` de `number` base, , et sont pris `boolean` en `void` `undefined` charge.</span><span class="sxs-lookup"><span data-stu-id="fa03d-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="fa03d-155">Les types Union utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="fa03d-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="fa03d-156">Les types de tableau sont autorisés s’ils sont de type `string` `number` , ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="fa03d-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="fa03d-157">Ils sont également autorisés si le type est une union prise en charge ou un type littéral pris en charge.</span><span class="sxs-lookup"><span data-stu-id="fa03d-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="fa03d-158">Les types d’objets utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="fa03d-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="fa03d-159">La saisie implicite est prise en charge, même si elle doit respecter les mêmes règles qu’un type défini.</span><span class="sxs-lookup"><span data-stu-id="fa03d-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="fa03d-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="fa03d-160">Example</span></span>

<span data-ttu-id="fa03d-161">La capture d’écran suivante montre un flux Power Automate qui est déclenché chaque fois [qu’un](https://github.com/) problème GitHub de sécurité vous est affecté.</span><span class="sxs-lookup"><span data-stu-id="fa03d-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="fa03d-162">Le flux exécute un script qui ajoute le problème à une table dans un Excel de travail.</span><span class="sxs-lookup"><span data-stu-id="fa03d-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="fa03d-163">S’il existe cinq problèmes ou plus dans ce tableau, le flux envoie un rappel par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="fa03d-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Éditeur Power Automate de flux affichant l’exemple de flux":::

<span data-ttu-id="fa03d-165">La fonction du script spécifie l’ID de problème et le titre du problème en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans le `main` tableau des problèmes.</span><span class="sxs-lookup"><span data-stu-id="fa03d-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="fa03d-166">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fa03d-166">See also</span></span>

- [<span data-ttu-id="fa03d-167">Exécuter Office scripts dans Excel sur le Web avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="fa03d-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="fa03d-168">Transmettre des données à des scripts dans un flux automatique Power Automate</span><span class="sxs-lookup"><span data-stu-id="fa03d-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="fa03d-169">Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement</span><span class="sxs-lookup"><span data-stu-id="fa03d-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="fa03d-170">Informations de dépannage pour les Power Automate avec Office scripts</span><span class="sxs-lookup"><span data-stu-id="fa03d-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="fa03d-171">Prise en main de Power Automate</span><span class="sxs-lookup"><span data-stu-id="fa03d-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="fa03d-172">Excel Documentation de référence sur le connecteur en ligne (Entreprise)</span><span class="sxs-lookup"><span data-stu-id="fa03d-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
