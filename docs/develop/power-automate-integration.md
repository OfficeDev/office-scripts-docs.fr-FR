---
title: Exécuter des scripts Office avec Power automate
description: Comment obtenir des scripts Office pour Excel sur le Web avec un flux de travail Automated Power.
ms.date: 07/24/2020
localization_priority: Normal
ms.openlocfilehash: 87bd4e15ef7680a7456077494e3fda8208d6b9d8
ms.sourcegitcommit: e9a8ef5f56177ea9a3d2fc5ac636368e5bdae1f4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/01/2020
ms.locfileid: "47321571"
---
# <a name="run-office-scripts-with-power-automate"></a>Exécuter des scripts Office avec Power automate

[Power automate](https://flow.microsoft.com) vous permet d’ajouter des scripts Office à un flux de travail automatisé plus important. Vous pouvez utiliser Power automate effectuer des opérations comme ajouter le contenu d’un message électronique à la table d’une feuille de calcul ou créer des actions dans vos outils de gestion de projet en fonction des commentaires de votre classeur.

## <a name="getting-started"></a>Prise en main

Si vous ne connaissez pas l’automate de puissance, nous vous recommandons de consulter la [prise en main de Power automate](/power-automate/getting-started). Ici, vous pouvez en savoir plus sur toutes les possibilités d’automatisation disponibles. Les documents ici se concentrent sur la façon dont les scripts Office fonctionnent avec automate d’alimentation et sur la façon d’améliorer votre expérience Excel.

Pour commencer à combiner les scripts Power Automated et Office, suivez le didacticiel [commencer à utiliser des scripts avec Power automate](../tutorials/excel-power-automate-manual.md). Cela vous apprend à créer un flux qui appelle un script simple. Une fois que vous avez terminé ce didacticiel et que vous avez [transmis des données à des scripts dans un didacticiel de mise à niveau automatique de l’alimentation automatique](../tutorials/excel-power-automate-trigger.md) , renvoyez ici pour obtenir des informations détaillées sur la connexion de scripts Office à la mise à niveau automatique des flux.

## <a name="excel-online-business-connector"></a>Connecteur Excel Online (Business)

Les [connecteurs](/connectors/connectors) sont les ponts entre l’automate de puissance et les applications. Le [connecteur Excel Online (Business)](/connectors/excelonlinebusiness) donne accès à vos flux aux classeurs Excel. L’action « exécuter un script » vous permet d’appeler n’importe quel script Office accessible via le classeur sélectionné. Vous pouvez également attribuer des paramètres d’entrée de scripts afin que les données puissent être fournies par le flux, ou que votre script renvoie des informations pour les étapes ultérieures dans le flux.

> [!IMPORTANT]
> L’action « exécuter un script » permet aux personnes qui utilisent le connecteur Excel d’accéder à votre classeur et à ses données. De plus, il existe des risques de sécurité pour les scripts qui effectuent des appels d’API externes, comme expliqué dans la rubrique [appels externes de Power Automated](external-calls.md). Si votre administrateur est concerné par l’exposition de données hautement sensibles, il peut soit désactiver le connecteur Excel Online, soit restreindre l’accès aux scripts Office via les contrôles de l' [administrateur des scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Transfert de données dans les flux pour les scripts

Power automate vous permet de transmettre des éléments de données entre les étapes de votre flux. Les scripts peuvent être configurés pour accepter tous les types d’informations dont vous avez besoin et renvoyer tout élément de votre classeur souhaité dans votre flux. L’entrée de votre script est spécifiée en ajoutant des paramètres à la `main` fonction (en plus de `workbook: ExcelScript.Workbook` ). La sortie du script est déclarée en ajoutant un type de retour à `main` .

> [!NOTE]
> Lorsque vous créez un bloc de script d’exécution dans votre flux, les paramètres acceptés et les types renvoyés sont renseignés. Si vous modifiez les paramètres ou les types de retour de votre script, vous devez rétablir le bloc de script « exécuter le script » de votre flux. Cela permet de s’assurer que les données sont analysées correctement.

Les sections suivantes couvrent les détails de l’entrée et de la sortie des scripts utilisés dans Power automate. Si vous souhaitez obtenir une approche pratique de l’apprentissage de cette rubrique, essayez de [transmettre des données à des scripts dans un didacticiel de puissance automate d’alimentation automatique](../tutorials/excel-power-automate-trigger.md) ou explorez le scénario d’exemple de [rappels de tâche automatisée](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-passing-data-to-a-script"></a>`main` Paramètres : transmission de données à un script

Toutes les entrées de script sont spécifiées comme paramètres supplémentaires pour la `main` fonction. Par exemple, si vous souhaitez qu’un script accepte un `string` qui représente un nom comme entrée, vous devez remplacer la `main` signature par `function main(workbook: ExcelScript.Workbook, name: string)` .

Lorsque vous configurez un flux dans Power Automated, vous pouvez spécifier des entrées de script sous forme de valeurs statiques, d' [expressions](/power-automate/use-expressions-in-conditions)ou de contenu dynamique. Pour plus d’informations sur le connecteur d’un service individuel, consultez la [documentation relative](/connectors/)à la mise à niveau automatique du connecteur.

Lors de l’ajout de paramètres d’entrée à la fonction d’un script `main` , tenez compte des quotas et des restrictions suivantes.

1. Le premier paramètre doit être de type `ExcelScript.Workbook` . Le nom de son paramètre n’a pas d’importance.

2. Chaque paramètre doit avoir un type (par exemple, `string` ou `number` ).

3. Les types de base,,,,, `string` `number` `boolean` `any` `unknown` `object` et `undefined` sont pris en charge.

4. Les tableaux des types de base précédemment répertoriés sont pris en charge.

5. Les tableaux imbriqués sont pris en charge en tant que paramètres (mais pas en tant que types de retour).

6. Les types Union sont autorisés s’ils sont une Union de littéraux appartenant à un même type (tel que `"Left" | "Right"` ). Les unions d’un type pris en charge avec undefined sont également prises en charge (par exemple, `string | undefined` ).

7. Les types d’objet sont autorisés s’ils contiennent des propriétés de type `string` ,,, des `number` `boolean` tableaux pris en charge ou d’autres objets pris en charge. L’exemple suivant montre les objets imbriqués pris en charge en tant que types de paramètres :

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

8. La définition de l’interface ou de la classe des objets doit être définie dans le script. Un objet peut également être défini de manière anonyme, comme dans l’exemple suivant :

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Les paramètres facultatifs sont autorisés et peuvent être dénotés comme tels à l’aide du modificateur facultatif `?` (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Les valeurs de paramètre par défaut sont autorisées (par exemple `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .

### <a name="returning-data-from-a-script"></a>Renvoi de données à partir d’un script

Les scripts peuvent renvoyer des données à partir du classeur afin d’être utilisées en tant que contenu dynamique dans un flux automatique de l’alimentation. Comme avec les paramètres d’entrée, Power automate place certaines restrictions sur le type de retour.

1. Les types de base,,, `string` `number` `boolean` `void` et `undefined` sont pris en charge.

2. Les types d’Union utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.

3. Les types tableau sont autorisés s’ils sont de type `string` , `number` ou `boolean` . Elles sont également autorisées si le type est un type de littéral Union pris en charge ou pris en charge.

4. Les types d’objets utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.

5. Le typage implicite est pris en charge, mais il doit suivre les mêmes règles qu’un type défini.

## <a name="avoid-using-relative-references"></a>Éviter d’utiliser des références relatives

Power automate exécute votre script dans le classeur Excel choisi de votre part. Le classeur peut être fermé lorsque cela se produit. Toutes les API qui s’appuient sur l’état actuel de l’utilisateur, telles que `Workbook.getActiveWorksheet` , échouent lorsqu’elles sont exécutées via Power Automated. Lors de la conception de vos scripts, veillez à utiliser des références absolues pour les feuilles de calcul et les plages.

Les méthodes suivantes génèrent une erreur et échouent lorsqu’elles sont appelées à partir d’un script dans un flux d’automate de puissance.

| Class | Méthode |
|--|--|
| [Graphique](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [Feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

## <a name="example"></a>Exemple

La capture d’écran suivante montre un flux automatique de puissance déclenché à chaque fois qu’un problème [GitHub](https://github.com/) vous est affecté. Le flux exécute un script qui ajoute le problème à un tableau dans un classeur Excel. Si ce tableau comporte au moins cinq problèmes, le flux envoie un rappel par courrier électronique.

![Exemple de flux tel qu’illustré dans l’éditeur de flux Automated Power.](../images/power-automate-parameter-return-sample.png)

La `main` fonction du script spécifie l’ID du problème et le titre du problème en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans la table des problèmes.

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

## <a name="see-also"></a>Voir aussi

- [Exécuter des scripts Office dans Excel sur le Web avec Power Automated Power](../tutorials/excel-power-automate-manual.md)
- [Transmettre des données à des scripts dans un flux automatique Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Principes de base pour la rédaction de scripts Office en Excel sur le web](scripting-fundamentals.md)
- [Prise en main de Power Automate](/power-automate/getting-started)
- [Documentation de référence du connecteur Excel Online (Business)](/connectors/excelonlinebusiness/)
