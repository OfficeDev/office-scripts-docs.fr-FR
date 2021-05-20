---
title: Exécutez Office scripts avec Power Automate
description: Comment obtenir des scripts Office pour Excel sur le Web avec un flux de travail Power Automate travail.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7562a2b2359cde67a9a47e0640515018fe23ac35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545039"
---
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="7c311-103">Exécutez Office scripts avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="7c311-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="7c311-104">[Power Automate vous](https://flow.microsoft.com) permet d’ajouter Office scripts à un flux de travail automatisé plus grand.</span><span class="sxs-lookup"><span data-stu-id="7c311-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="7c311-105">Vous pouvez utiliser Power Automate des choses comme ajouter le contenu d’un e-mail à la table d’une feuille de travail ou créer des actions dans vos outils de gestion de projet s’appuyant sur des commentaires de manuel.</span><span class="sxs-lookup"><span data-stu-id="7c311-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span>

## <a name="get-started"></a><span data-ttu-id="7c311-106">Prise en main</span><span class="sxs-lookup"><span data-stu-id="7c311-106">Get started</span></span>

<span data-ttu-id="7c311-107">Si vous êtes nouveau à Power Automate, nous vous recommandons de [visiter Démarrer avec Power Automate](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="7c311-107">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="7c311-108">Là, vous pouvez en apprendre davantage sur toutes les possibilités d’automatisation à votre disposition.</span><span class="sxs-lookup"><span data-stu-id="7c311-108">There, you can learn more about all the automation possibilities available to you.</span></span> <span data-ttu-id="7c311-109">Les documents ici se concentrent sur la façon dont Office scripts fonctionnent avec Power Automate et comment cela peut aider à améliorer votre expérience Excel vie.</span><span class="sxs-lookup"><span data-stu-id="7c311-109">The documents here focus on how Office Scripts work with Power Automate and how that can help improve your Excel experience.</span></span>

<span data-ttu-id="7c311-110">Pour commencer à combiner Power Automate scripts Office, suivez le tutoriel [Commencez à utiliser des scripts avec Power Automate](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="7c311-110">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="7c311-111">Cela vous apprendra à créer un flux qui appelle un script simple.</span><span class="sxs-lookup"><span data-stu-id="7c311-111">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="7c311-112">Une fois que vous avez terminé ce tutoriel et [les données Pass aux scripts dans un tutoriel de flux de Power Automate exécuté](../tutorials/excel-power-automate-trigger.md) automatiquement, revenez ici pour plus d’informations détaillées sur la connexion des scripts Office aux flux Power Automate utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="7c311-112">After you've completed that tutorial and the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="7c311-113">Excel Connecteur en ligne (Business)</span><span class="sxs-lookup"><span data-stu-id="7c311-113">Excel Online (Business) connector</span></span>

<span data-ttu-id="7c311-114">[Les connecteurs](/connectors/connectors) sont les ponts entre Power Automate et les applications.</span><span class="sxs-lookup"><span data-stu-id="7c311-114">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="7c311-115">Le [Excel connecteur en ligne (Business)](/connectors/excelonlinebusiness) donne à vos flux accès à Excel manuels.</span><span class="sxs-lookup"><span data-stu-id="7c311-115">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="7c311-116">L’action « Exécuter le script » vous permet d’appeler n’importe Office script accessible via le cahier de travail sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7c311-116">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="7c311-117">Vous pouvez également donner à vos scripts des paramètres d’entrée afin que les données puissent être fournies par le flux, ou avoir vos informations de retour de script pour les étapes ultérieures du flux.</span><span class="sxs-lookup"><span data-stu-id="7c311-117">You can also give your scripts input parameters so data can be provided by the flow, or have your script return information for later steps in the flow.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7c311-118">L’action « Exécuter script » donne aux personnes qui utilisent le connecteur Excel un accès significatif à votre cahier de travail et à ses données.</span><span class="sxs-lookup"><span data-stu-id="7c311-118">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="7c311-119">En outre, il ya des risques de sécurité avec les scripts qui font des appels API externes, comme expliqué [dans les appels externes de Power Automate](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="7c311-119">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="7c311-120">Si votre administrateur est préoccupé par l’exposition de données hautement sensibles, ils peuvent soit désactiver le connecteur Excel Online ou restreindre l’accès aux scripts Office par le [biais des contrôles de l’administrateur scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="7c311-120">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="data-transfer-in-flows-for-scripts"></a><span data-ttu-id="7c311-121">Transfert de données dans les flux de scripts</span><span class="sxs-lookup"><span data-stu-id="7c311-121">Data transfer in flows for scripts</span></span>

<span data-ttu-id="7c311-122">Power Automate vous permet de transmettre des éléments de données entre les étapes de votre flux.</span><span class="sxs-lookup"><span data-stu-id="7c311-122">Power Automate lets you pass pieces of data between steps of your flow.</span></span> <span data-ttu-id="7c311-123">Les scripts peuvent être configurés pour accepter tous les types d’informations dont vous avez besoin et renvoyer tout ce qui vient de votre cahier de travail que vous souhaitez dans votre flux.</span><span class="sxs-lookup"><span data-stu-id="7c311-123">Scripts can be configured to accept whatever types of information you need and return anything from your workbook that you want in your flow.</span></span> <span data-ttu-id="7c311-124">L’entrée de votre script est spécifiée en ajoutant des paramètres à `main` la fonction (en plus de `workbook: ExcelScript.Workbook` ).</span><span class="sxs-lookup"><span data-stu-id="7c311-124">Input for your script is specified by adding parameters to the `main` function (in addition to `workbook: ExcelScript.Workbook`).</span></span> <span data-ttu-id="7c311-125">La sortie du script est déclarée en ajoutant un type de retour à `main` .</span><span class="sxs-lookup"><span data-stu-id="7c311-125">Output from the script is declared by adding a return type to `main`.</span></span>

> [!NOTE]
> <span data-ttu-id="7c311-126">Lorsque vous créez un bloc « Script d’exécuter » dans votre flux, les paramètres acceptés et les types retournés sont remplis.</span><span class="sxs-lookup"><span data-stu-id="7c311-126">When you create a "Run Script" block in your flow, the accepted parameters and returned types are populated.</span></span> <span data-ttu-id="7c311-127">Si vous modifiez les paramètres ou retournez les types de votre script, vous devrez refaire le bloc « Exécuter le script » de votre flux.</span><span class="sxs-lookup"><span data-stu-id="7c311-127">If you change the parameters or return types of your script, you'll need to redo the "Run script" block of your flow.</span></span> <span data-ttu-id="7c311-128">Cela garantit que les données sont correctement analyses.</span><span class="sxs-lookup"><span data-stu-id="7c311-128">This ensures the data is being parsed correctly.</span></span>

<span data-ttu-id="7c311-129">Les sections suivantes couvrent les détails de l’entrée et de la sortie des scripts utilisés dans Power Automate.</span><span class="sxs-lookup"><span data-stu-id="7c311-129">The following sections cover the details of input and output for scripts used in Power Automate.</span></span> <span data-ttu-id="7c311-130">Si vous souhaitez une approche pratique pour l’apprentissage de ce sujet, essayez les données Pass aux scripts dans un [didacticiel de flux de Power Automate exécuté](../tutorials/excel-power-automate-trigger.md) automatiquement ou explorez le scénario d’exemple de [rappels de tâches automatisés.](../resources/scenarios/task-reminders.md)</span><span class="sxs-lookup"><span data-stu-id="7c311-130">If you'd like a hands-on approach to learning this topic, try out the [Pass data to scripts in an automatically-run Power Automate flow](../tutorials/excel-power-automate-trigger.md) tutorial or explore the [Automated task reminders](../resources/scenarios/task-reminders.md) sample scenario.</span></span>

### <a name="main-parameters-pass-data-to-a-script"></a><span data-ttu-id="7c311-131">`main` Paramètres : Transmettre des données à un script</span><span class="sxs-lookup"><span data-stu-id="7c311-131">`main` Parameters: Pass data to a script</span></span>

<span data-ttu-id="7c311-132">Toutes les entrées de script sont spécifiées comme paramètres supplémentaires pour la `main` fonction.</span><span class="sxs-lookup"><span data-stu-id="7c311-132">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="7c311-133">Par exemple, si vous vouliez qu’un script accepte un `string` nom qui représente un nom comme entrée, vous changeriez la signature en `main` `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="7c311-133">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="7c311-134">Lorsque vous configurez un flux dans les Power Automate, vous pouvez spécifier l’entrée du script sous forme de valeurs [statiques, d’expressions](/power-automate/use-expressions-in-conditions)ou de contenu dynamique.</span><span class="sxs-lookup"><span data-stu-id="7c311-134">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="7c311-135">Les détails sur le connecteur d’un service individuel peuvent être trouvés [dans la documentation Power Automate Connecteur.](/connectors/)</span><span class="sxs-lookup"><span data-stu-id="7c311-135">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="7c311-136">Lorsque vous ajoutez des paramètres d’entrée à la fonction `main` d’un script, considérez les allocations et restrictions suivantes.</span><span class="sxs-lookup"><span data-stu-id="7c311-136">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="7c311-137">Le premier paramètre doit être de type `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="7c311-137">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="7c311-138">Son nom de paramètre n’a pas d’importance.</span><span class="sxs-lookup"><span data-stu-id="7c311-138">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="7c311-139">Chaque paramètre doit avoir un type (tel `string` ou `number` ).</span><span class="sxs-lookup"><span data-stu-id="7c311-139">Every parameter must have a type (such as `string` or `number`).</span></span>

3. <span data-ttu-id="7c311-140">Les types de `string` base , , , , , et sont pris en `number` `boolean` `unknown` `object` `undefined` charge.</span><span class="sxs-lookup"><span data-stu-id="7c311-140">The basic types `string`, `number`, `boolean`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="7c311-141">Les tableaux des types de base précédemment répertoriés sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="7c311-141">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="7c311-142">Les tableaux imbriqués sont pris en charge sous forme de paramètres (mais pas en tant que types de retour).</span><span class="sxs-lookup"><span data-stu-id="7c311-142">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="7c311-143">Les types d’union sont autorisés s’il s’agit d’une union de littérales appartenant à un seul type `"Left" | "Right"` (comme).</span><span class="sxs-lookup"><span data-stu-id="7c311-143">Union types are allowed if they are a union of literals belonging to a single type (such as `"Left" | "Right"`).</span></span> <span data-ttu-id="7c311-144">Les unions d’un type soutenu avec des non définis sont également soutenues (telles que `string | undefined` ).</span><span class="sxs-lookup"><span data-stu-id="7c311-144">Unions of a supported type with undefined are also supported (such as `string | undefined`).</span></span>

7. <span data-ttu-id="7c311-145">Les types d’objets sont autorisés s’ils contiennent des propriétés de `string` `number` type, `boolean` , des tableaux pris en charge, ou d’autres objets pris en charge.</span><span class="sxs-lookup"><span data-stu-id="7c311-145">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="7c311-146">L’exemple suivant montre les objets imbriqués qui sont pris en charge sous forme de types de paramètres :</span><span class="sxs-lookup"><span data-stu-id="7c311-146">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="7c311-147">Les objets doivent avoir leur interface ou définition de classe définie dans le script.</span><span class="sxs-lookup"><span data-stu-id="7c311-147">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="7c311-148">Un objet peut également être défini anonymement en ligne, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="7c311-148">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="7c311-149">Les paramètres optionnels sont autorisés et peuvent être indiqués comme tels en utilisant le modificateur `?` optionnel (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="7c311-149">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="7c311-150">Les valeurs de paramètres par défaut sont autorisées (par exemple `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="7c311-150">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

### <a name="return-data-from-a-script"></a><span data-ttu-id="7c311-151">Renvoyer les données d’un script</span><span class="sxs-lookup"><span data-stu-id="7c311-151">Return data from a script</span></span>

<span data-ttu-id="7c311-152">Les scripts peuvent renvoyer des données du cahier de travail pour les utiliser comme contenu dynamique dans un flux Power Automate fluide.</span><span class="sxs-lookup"><span data-stu-id="7c311-152">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="7c311-153">Comme pour les paramètres d’entrée, Power Automate impose certaines restrictions sur le type de retour.</span><span class="sxs-lookup"><span data-stu-id="7c311-153">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="7c311-154">Les types de base `string` , , , et sont pris en `number` `boolean` `void` `undefined` charge.</span><span class="sxs-lookup"><span data-stu-id="7c311-154">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="7c311-155">Les types d’union utilisés comme types de retour suivent les mêmes restrictions qu’ils le font lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="7c311-155">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="7c311-156">Les types de tableaux sont autorisés s’ils sont de `string` `number` type, ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="7c311-156">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="7c311-157">Ils sont également autorisés si le type est un syndicat soutenu ou soutenu type littéral.</span><span class="sxs-lookup"><span data-stu-id="7c311-157">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="7c311-158">Les types d’objets utilisés comme types de retour suivent les mêmes restrictions qu’ils le font lorsqu’ils sont utilisés comme paramètres de script.</span><span class="sxs-lookup"><span data-stu-id="7c311-158">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="7c311-159">La dactylographie implicite est prise en charge, bien qu’elle doit suivre les mêmes règles qu’un type défini.</span><span class="sxs-lookup"><span data-stu-id="7c311-159">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="example"></a><span data-ttu-id="7c311-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="7c311-160">Example</span></span>

<span data-ttu-id="7c311-161">La capture d’écran suivante Power Automate un flux de flux qui est déclenché chaque [fois qu’GitHub](https://github.com/) problème est attribué à vous.</span><span class="sxs-lookup"><span data-stu-id="7c311-161">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="7c311-162">Le flux exécute un script qui ajoute le problème à une table dans un Excel de travail.</span><span class="sxs-lookup"><span data-stu-id="7c311-162">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="7c311-163">S’il y a cinq problèmes ou plus dans ce tableau, le flux envoie un rappel par courriel.</span><span class="sxs-lookup"><span data-stu-id="7c311-163">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="L’éditeur Power Automate débit de flux montrant le flux d’exemple":::

<span data-ttu-id="7c311-165">La `main` fonction du script spécifie l’ID d’émission et le titre d’émission en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans le tableau d’émission.</span><span class="sxs-lookup"><span data-stu-id="7c311-165">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="7c311-166">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7c311-166">See also</span></span>

- [<span data-ttu-id="7c311-167">Exécutez Office scripts en Excel sur le Web avec Power Automate</span><span class="sxs-lookup"><span data-stu-id="7c311-167">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="7c311-168">Transmettre des données à des scripts dans un flux automatique Power Automate</span><span class="sxs-lookup"><span data-stu-id="7c311-168">Pass data to scripts in an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="7c311-169">Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement</span><span class="sxs-lookup"><span data-stu-id="7c311-169">Return data from a script to an automatically-run Power Automate flow</span></span>](../tutorials/excel-power-automate-returns.md)
- [<span data-ttu-id="7c311-170">Informations de dépannage pour Power Automate avec Office scripts</span><span class="sxs-lookup"><span data-stu-id="7c311-170">Troubleshooting information for Power Automate with Office Scripts</span></span>](../testing/power-automate-troubleshooting.md)
- [<span data-ttu-id="7c311-171">Prise en main de Power Automate</span><span class="sxs-lookup"><span data-stu-id="7c311-171">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="7c311-172">Excel Documentation de référence connecteur en ligne (Business)</span><span class="sxs-lookup"><span data-stu-id="7c311-172">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
