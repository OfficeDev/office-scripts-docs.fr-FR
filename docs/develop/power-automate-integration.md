---
title: Exécuter Office scripts avec Power Automate
description: Comment obtenir des scripts Office pour Excel sur le Web un flux de travail Power Automate de travail.
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: f7358b79248974ddb548b54437422670a37531bf
ms.sourcegitcommit: 79ce4fad6d284b1aa71f5ad6d2938d9ad6a09fee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/12/2022
ms.locfileid: "63459619"
---
# <a name="run-office-scripts-with-power-automate"></a>Exécuter Office scripts avec Power Automate

[Power Automate](https://flow.microsoft.com) vous permet d’ajouter Office scripts à un flux de travail automatisé plus important. Vous pouvez utiliser Power Automate des opérations telles que l’ajout du contenu d’un e-mail au tableau d’une feuille de calcul ou la création d’actions dans vos outils de gestion de projet en fonction des commentaires de votre feuille de calcul.

## <a name="get-started"></a>Prise en main

Si vous débutez avec Power Automate, nous vous recommandons de consulter La mise en [Power Automate.](/power-automate/getting-started) Vous y découvrirez toutes les possibilités d’automatisation disponibles. Les documents présentés ici se concentrent sur Office l’Power Automate scripts et sur la façon dont cela peut vous aider à améliorer Excel expérience utilisateur.

Pour commencer à combiner Power Automate et Office scripts, suivez le didacticiel Commencer à utiliser des [scripts avec Power Automate](../tutorials/excel-power-automate-manual.md). Cela vous montre comment créer un flux qui appelle un script simple. Une fois que vous avez terminé ce didacticiel et passé les données aux [scripts](../tutorials/excel-power-automate-trigger.md) dans un didacticiel de flux Power Automate exécuté automatiquement, revenir ici pour obtenir des informations détaillées sur la connexion de scripts Office à des flux Power Automate.

## <a name="excel-online-business-connector"></a>Excel Online (Entreprise)

[Les connecteurs](/connectors/connectors) sont les ponts entre Power Automate applications. Le [connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness) permet à vos flux d’accéder à Excel de travail. L’action « Exécuter le script » vous permet d’appeler Office script accessible via le livre de travail sélectionné. Vous pouvez également donner à vos scripts des paramètres d’entrée afin que les données soient fournies par le flux ou que votre script retourne des informations pour les étapes ultérieures du flux.

> [!IMPORTANT]
> L’action « Exécuter le script » donne aux personnes qui utilisent le connecteur Excel un accès significatif à votre workbook et à ses données. En outre, il existe des risques de sécurité avec les scripts qui appellent des API externes, comme expliqué dans les appels externes de [Power Automate](external-calls.md). Si votre administrateur est préoccupé par l’exposition de données hautement sensibles, il peut désactiver le connecteur Excel Online ou restreindre l’accès aux scripts Office par le biais des contrôles d’administrateur [Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Transfert de données dans les flux pour les scripts

Power Automate vous permet de passer des éléments de données entre les étapes de votre flux. Les scripts peuvent être configurés pour accepter les types d’informations dont vous avez besoin et renvoyer tout ce dont vous avez besoin dans votre flux de travail. L’entrée de votre script est spécifiée en ajoutant des paramètres à la `main` fonction (en plus de `workbook: ExcelScript.Workbook`). La sortie du script est déclarée en ajoutant un type de retour à `main`.

> [!NOTE]
> Lorsque vous créez un bloc « Exécuter un script » dans votre flux, les paramètres acceptés et les types renvoyés sont remplis. Si vous modifiez les paramètres ou renvoyez des types de votre script, vous devez redéfaire le bloc « Exécuter le script » de votre flux. Cela garantit que les données sont en cours d’analyse correctement.

Les sections suivantes couvrent les détails de l’entrée et de la sortie pour les scripts utilisés dans Power Automate. Si vous souhaitez une approche pratique de l’apprentissage de cette rubrique, essayez de transmettre des données aux [scripts](../tutorials/excel-power-automate-trigger.md) dans un didacticiel de flux Power Automate exécuté automatiquement ou explorez l’exemple de scénario de [rappels](../resources/scenarios/task-reminders.md) de tâches automatisés.

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Paramètres : transmettre des données à un script

Toutes les entrées de script sont spécifiées en tant que paramètres supplémentaires pour la `main` fonction. Par exemple, si vous souhaitez qu’un script accepte un `string` nom qui représente un nom comme entrée, vous devez modifier la `main` signature en `function main(workbook: ExcelScript.Workbook, name: string)`.

Lorsque vous configurez un flux dans Power Automate, vous pouvez spécifier une entrée de script en tant que valeurs statiques, [expressions](/power-automate/use-expressions-in-conditions) ou contenu dynamique. Pour plus d’informations sur le connecteur d’un service individuel, voir la [documentation Power Automate Connector](/connectors/).

#### <a name="type-restrictions"></a>Restrictions de type

Lorsque vous ajoutez des paramètres d’entrée à la fonction d’un `main` script, prenons en compte les restrictions et les allocations suivantes. Ceux-ci s’appliquent également au type de retour du script.

1. Le premier paramètre doit être de type `ExcelScript.Workbook`. Son nom de paramètre n’a pas d’importance.

1. Les types `string`, `number`, `boolean`, `unknown`et `object`sont `undefined` pris en charge.

1. Les tableaux (à la fois `[]` et `Array<T>` les styles) des types répertoriés précédemment sont pris en charge. Les tableaux imbrmbrés sont également pris en charge.

1. Les types Union sont autorisés s’il s’agit d’une union de littéraux appartenant à un seul type (par `"Left" | "Right"`exemple, et non `"Left", 5`). Les personnes d’un type pris en charge avec undefined sont également pris en charge (par exemple).`string | undefined`

1. Les types d’objets sont autorisés s’ils contiennent des propriétés de type `string``number`, , tableaux `boolean`pris en charge ou autres objets pris en charge. L’exemple suivant montre les objets imbrmbrés pris en charge en tant que types de paramètres.

    ```TypeScript
    // The Employee object is supported because Position is also composed of supported types.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

1. L’interface ou la définition de classe des objets doit être définie dans le script. Un objet peut également être défini de manière inline anonyme, comme dans l’exemple suivant.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>Paramètres facultatifs et par défaut

1. Les paramètres facultatifs sont autorisés et sont indiqués avec le `?` modificateur facultatif (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)`).

1. Les valeurs de paramètre par défaut sont autorisées (par exemple `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.

### <a name="return-data-from-a-script"></a>Renvoyer des données à partir d’un script

Les scripts peuvent renvoyer des données à partir du workbook à utiliser en tant que contenu dynamique dans Power Automate flux. Les [mêmes restrictions de type répertoriées précédemment](#type-restrictions) s’appliquent au type de retour. Pour renvoyer un objet, ajoutez la syntaxe de type de retour à la `main` fonction. Par exemple, si vous souhaitez renvoyer une `string` valeur à partir du script, votre `main` signature sera .`function main(workbook: ExcelScript.Workbook): string`

## <a name="example"></a>Exemple

La capture d’écran suivante montre Power Automate flux de données qui est déclenché chaque fois [qu’un GitHub](https://github.com/) est affecté. Le flux exécute un script qui ajoute le problème à une table dans un Excel de travail. S’il existe cinq problèmes ou plus dans ce tableau, le flux envoie un rappel par courrier électronique.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="L Power Automate de flux affichant l’exemple de flux.":::

La `main` fonction du script spécifie l’ID de problème et le titre du problème en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans le tableau des problèmes.

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

- [Exécuter Office scripts dans Excel sur le Web avec Power Automate](../tutorials/excel-power-automate-manual.md)
- [Transmettre des données à des scripts dans un flux automatique Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement](../tutorials/excel-power-automate-returns.md)
- [Informations de dépannage pour les Power Automate avec Office scripts](../testing/power-automate-troubleshooting.md)
- [Prise en main de Power Automate](/power-automate/getting-started)
- [Documentation de référence Excel Online (Business)](/connectors/excelonlinebusiness/)
