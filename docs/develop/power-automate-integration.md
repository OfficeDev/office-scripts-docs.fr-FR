---
title: Exécuter des scripts Office avec Power automate
description: Comment obtenir des scripts Office pour Excel sur le Web avec un flux de travail Automated Power.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 0ea58324998d23020e04cb37dfeea065791757f5
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043383"
---
# <a name="run-office-scripts-with-power-automate"></a>Exécuter des scripts Office avec Power automate

[Power automate](https://flow.microsoft.com) vous permet d’ajouter des scripts Office à un flux de travail automatisé plus important. Vous pouvez utiliser Power automate effectuer des opérations comme ajouter le contenu d’un message électronique à la table d’une feuille de calcul ou créer des actions dans vos outils de gestion de projet en fonction des commentaires de votre classeur. Si vous ne connaissez pas l’automate de puissance, nous vous recommandons de consulter la [prise en main de Power automate](/power-automate/getting-started). Ici, vous pouvez en savoir plus sur l’automatisation de vos flux de travail sur plusieurs services.

> [!IMPORTANT]
> Actuellement, vous ne pouvez pas exécuter des scripts Office à partir d’un [flux partagé](/power-automate/share-buttons). Seul l’utilisateur qui a créé un script peut l’exécuter, même via automate d’alimentation.

## <a name="getting-started"></a>Prise en main

Pour commencer à combiner les scripts Power Automated et Office, suivez le didacticiel [commencer à utiliser des scripts avec Power automate](../tutorials/excel-power-automate-manual.md). Cela vous apprend à créer un flux qui appelle un script simple. Une fois que vous avez terminé ce didacticiel et que vous avez [exécuté automatiquement les scripts avec Power Automated](../tutorials/excel-power-automate-trigger.md) , renvoyez ici pour obtenir des informations détaillées sur la connexion des scripts Office à la mise à niveau automatique des flux.

## <a name="excel-online-business-connector"></a>Connecteur Excel Online (Business)

Les [connecteurs](/connectors/connectors) sont les ponts entre l’automate de puissance et les applications. Le [connecteur Excel Online (Business)](/connectors/excelonlinebusiness) donne accès à vos flux aux classeurs Excel. L’action « exécuter un script » vous permet d’appeler n’importe quel script Office accessible via le classeur sélectionné. Vous pouvez non seulement exécuter des scripts via un flux, mais vous pouvez transmettre des données vers et depuis le classeur avec le flux via les scripts.

> [!IMPORTANT]
> L’action « exécuter un script » permet aux personnes qui utilisent le connecteur Excel d’accéder à votre classeur et à ses données. De plus, il existe des risques de sécurité pour les scripts qui effectuent des appels d’API externes, comme expliqué dans la rubrique [appels externes de Power Automated](external-calls.md). Si votre administrateur est concerné par l’exposition de données hautement sensibles, il peut soit désactiver le connecteur Excel Online, soit restreindre l’accès aux scripts Office via les contrôles de l' [administrateur des scripts Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).

## <a name="passing-data-from-power-automate-into-a-script"></a>Transmission de données de Power automate dans un script

Toutes les entrées de script sont spécifiées comme paramètres supplémentaires pour la `main` fonction. Par exemple, si vous souhaitez qu’un script accepte un `string` qui représente un nom comme entrée, vous devez remplacer la `main` signature par `function main(workbook: ExcelScript.Workbook, name: string)` .

Lorsque vous configurez un flux dans Power Automated, vous pouvez spécifier des entrées de script sous forme de valeurs statiques, d' [expressions](/power-automate/use-expressions-in-conditions)ou de contenu dynamique. Pour plus d’informations sur le connecteur d’un service individuel, consultez la [documentation relative](/connectors/)à la mise à niveau automatique du connecteur.

Lors de l’ajout de paramètres d’entrée à la fonction d’un script `main` , tenez compte des quotas et des restrictions suivantes.

1. Le premier paramètre doit être de type `ExcelScript.Workbook` . Le nom de son paramètre n’a pas d’importance.

2. Chaque paramètre doit avoir un type.

3. Les types de base,,,,, `string` `number` `boolean` `any` `unknown` `object` et `undefined` sont pris en charge.

4. Les tableaux des types de base précédemment répertoriés sont pris en charge.

5. Les tableaux imbriqués sont pris en charge en tant que paramètres (mais pas en tant que types de retour).

6. Les types Union sont autorisés s’il s’agit d’une Union de littéraux appartenant à un seul type ( `string` , `number` , ou `boolean` ). Les unions d’un type pris en charge avec undefined sont également prises en charge.

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

## <a name="returning-data-from-a-script-back-to-power-automate"></a>Retour des données d’un script à automate d’alimentation

Les scripts peuvent renvoyer des données à partir du classeur afin d’être utilisées en tant que contenu dynamique dans un flux automatique de l’alimentation. Comme avec les paramètres d’entrée, Power automate place certaines restrictions sur le type de retour.

1. Les types de base,,, `string` `number` `boolean` `void` et `undefined` sont pris en charge.

2. Les types d’Union utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.

3. Les types tableau sont autorisés s’ils sont de type `string` , `number` ou `boolean` . Elles sont également autorisées si le type est un type de littéral Union pris en charge ou pris en charge.

4. Les types d’objets utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.

5. Le typage implicite est pris en charge, mais il doit suivre les mêmes règles qu’un type défini.

## <a name="avoid-using-relative-references"></a>Éviter d’utiliser des références relatives

Power automate exécute votre script dans le classeur Excel choisi de votre part. Le classeur peut être fermé lorsque cela se produit. Toutes les API qui s’appuient sur l’état actuel de l’utilisateur, telles que `Workbook.getActiveWorksheet` , échouent lorsqu’elles sont exécutées via Power Automated. Lors de la conception de vos scripts, veillez à utiliser des références absolues pour les feuilles de calcul et les plages.

Les fonctions suivantes génèrent une erreur et échouent lorsqu’elles sont appelées à partir d’un script dans un flux d’automate de puissance.

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
- [Exécuter automatiquement des scripts avec Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Principes de base des scripts pour Office Scripts dans Excel sur le web](scripting-fundamentals.md)
- [Prise en main de Power Automate](/power-automate/getting-started)
- [Documentation de référence du connecteur Excel Online (Business)](/connectors/excelonlinebusiness/)
