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
# <a name="run-office-scripts-with-power-automate"></a>Exécuter Office scripts avec Power Automate

[Power Automate](https://flow.microsoft.com) vous permet d’ajouter Office scripts à un flux de travail automatisé plus important. Vous pouvez utiliser Power Automate des opérations telles que l’ajout du contenu d’un e-mail au tableau d’une feuille de calcul ou la création d’actions dans vos outils de gestion de projet en fonction des commentaires de votre feuille de calcul.

## <a name="getting-started"></a>Prise en main

Si vous débutez avec Power Automate, nous vous recommandons de visiter La mise en [Power Automate](/power-automate/getting-started). Vous y découvrirez toutes les possibilités d’automatisation disponibles. Les documents présentés ici se concentrent sur Office fonctionnement des scripts Power Automate et sur la façon dont cela peut vous aider à améliorer Excel expérience utilisateur.

Pour commencer à combiner Power Automate et Office scripts, suivez le didacticiel Démarrer à l’aide de [scripts Power Automate](../tutorials/excel-power-automate-manual.md). Cela vous montre comment créer un flux qui appelle un script simple. Une fois que vous avez terminé ce didacticiel et passé les données aux [scripts](../tutorials/excel-power-automate-trigger.md) dans un didacticiel de flux Power Automate exécuté automatiquement, revenir ici pour obtenir des informations détaillées sur la connexion de scripts Office à des flux Power Automate.

## <a name="excel-online-business-connector"></a>Excel Connecteur En ligne (Entreprise)

[Les connecteurs](/connectors/connectors) sont les ponts entre Power Automate applications. Le [connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness) permet à vos flux d’accéder Excel de travail. L’action « Exécuter le script » vous permet d’appeler Office script accessible via le livre de travail sélectionné. Vous pouvez également donner à vos scripts des paramètres d’entrée afin que les données soient fournies par le flux ou que votre script retourne des informations pour les étapes ultérieures du flux.

> [!IMPORTANT]
> L’action « Exécuter le script » donne aux utilisateurs du connecteur Excel un accès significatif à votre workbook et à ses données. En outre, il existe des risques de sécurité avec les scripts qui font des appels d’API externes, comme expliqué dans les appels externes de [Power Automate](external-calls.md). Si votre administrateur est préoccupé par l’exposition de données hautement sensibles, il peut désactiver le connecteur Excel Online ou restreindre l’accès aux scripts Office par le biais des contrôles d’administrateur [Office Scripts](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Transfert de données dans les flux pour les scripts

Power Automate vous permet de passer des éléments de données entre les étapes de votre flux. Les scripts peuvent être configurés pour accepter les types d’informations dont vous avez besoin et renvoyer tout ce dont vous avez besoin dans votre flux de travail. L’entrée de votre script est spécifiée en ajoutant des paramètres à la `main` fonction (en plus de `workbook: ExcelScript.Workbook` ). La sortie du script est déclarée en ajoutant un type de retour à `main` .

> [!NOTE]
> Lorsque vous créez un bloc « Exécuter un script » dans votre flux, les paramètres acceptés et les types renvoyés sont remplis. Si vous modifiez les paramètres ou renvoyez des types de votre script, vous devrez revenir au bloc « Exécuter le script » de votre flux. Cela garantit que les données sont en cours d’analyse correctement.

Les sections suivantes couvrent les détails de l’entrée et de la sortie pour les scripts utilisés dans Power Automate. Si vous souhaitez une approche pratique de l’apprentissage de cette rubrique, essayez de transmettre des données aux [scripts](../tutorials/excel-power-automate-trigger.md) dans un didacticiel de flux Power Automate exécuté automatiquement ou explorez l’exemple de scénario de [rappels](../resources/scenarios/task-reminders.md) de tâches automatisés.

### <a name="main-parameters-passing-data-to-a-script"></a>`main` Paramètres : transmission de données à un script

Toutes les entrées de script sont spécifiées en tant que paramètres supplémentaires pour la `main` fonction. Par exemple, si vous souhaitez qu’un script accepte un nom qui représente un nom comme entrée, vous devez modifier `string` la `main` signature en `function main(workbook: ExcelScript.Workbook, name: string)` .

Lorsque vous configurez un flux dans Power Automate, vous pouvez spécifier une entrée de script en tant que valeurs statiques, [expressions](/power-automate/use-expressions-in-conditions)ou contenu dynamique. Pour plus d’informations sur le connecteur d’un service individuel, voir la [documentation Power Automate Connector.](/connectors/)

Lorsque vous ajoutez des paramètres d’entrée à la fonction d’un script, prenons en compte les `main` limites et restrictions suivantes.

1. Le premier paramètre doit être de type `ExcelScript.Workbook` . Son nom de paramètre n’a pas d’importance.

2. Chaque paramètre doit avoir un type (par `string` exemple, ou `number` ).

3. Les types `string` de base , , , , et sont pris en `number` `boolean` `any` `unknown` `object` `undefined` charge.

4. Les tableaux des types de base répertoriés précédemment sont pris en charge.

5. Les tableaux imbrmbrés sont pris en charge en tant que paramètres (mais pas en tant que types de retour).

6. Les types Union sont autorisés s’il s’agit d’une union de littéraux appartenant à un seul type (par `"Left" | "Right"` exemple). Les personnes d’un type pris en charge avec undefined sont également pris en charge (par `string | undefined` exemple).

7. Les types d’objets sont autorisés s’ils contiennent des propriétés de type , , tableaux pris `string` en charge ou autres objets pris en `number` `boolean` charge. L’exemple suivant montre les objets imbrmbrés pris en charge en tant que types de paramètres :

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

8. L’interface ou la définition de classe des objets doit être définie dans le script. Un objet peut également être défini de manière anonyme en ligne, comme dans l’exemple suivant :

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Les paramètres facultatifs sont autorisés et peuvent être indiqués en tant que tels à l’aide du modificateur facultatif `?` (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Les valeurs de paramètre par défaut sont autorisées (par `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` exemple.

### <a name="returning-data-from-a-script"></a>Renvoi de données à partir d’un script

Les scripts peuvent renvoyer des données à partir du workbook à utiliser en tant que contenu dynamique dans Power Automate flux. Comme pour les paramètres d’entrée, Power Automate des restrictions sur le type de retour.

1. Les types `string` de `number` base, , et sont pris `boolean` en `void` `undefined` charge.

2. Les types Union utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.

3. Les types de tableau sont autorisés s’ils sont de type `string` `number` , ou `boolean` . Ils sont également autorisés si le type est une union prise en charge ou un type littéral pris en charge.

4. Les types d’objets utilisés comme types de retour suivent les mêmes restrictions que lorsqu’ils sont utilisés comme paramètres de script.

5. La saisie implicite est prise en charge, même si elle doit respecter les mêmes règles qu’un type défini.

## <a name="example"></a>Exemple

La capture d’écran suivante montre un flux Power Automate qui est déclenché chaque fois [qu’un](https://github.com/) problème GitHub de sécurité vous est affecté. Le flux exécute un script qui ajoute le problème à une table dans un Excel de travail. S’il existe cinq problèmes ou plus dans ce tableau, le flux envoie un rappel par courrier électronique.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Éditeur Power Automate de flux affichant l’exemple de flux":::

La fonction du script spécifie l’ID de problème et le titre du problème en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans le `main` tableau des problèmes.

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
- [Excel Documentation de référence sur le connecteur en ligne (Entreprise)](/connectors/excelonlinebusiness/)
