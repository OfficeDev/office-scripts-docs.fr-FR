---
title: Exécuter des scripts Office avec Power Automate
description: Comment obtenir des scripts Office pour Excel sur le Web utiliser un workflow Power Automate.
ms.date: 06/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 61e51861bd2c987c25d40e9ac6d2247122256918
ms.sourcegitcommit: c5ffe0a95b962936ee92e7ffe17388bef6d4fad8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/29/2022
ms.locfileid: "66241860"
---
# <a name="run-office-scripts-with-power-automate"></a>Exécuter des scripts Office avec Power Automate

[Power Automate](https://flow.microsoft.com) vous permet d’ajouter des scripts Office à un workflow plus volumineux et automatisé. Vous pouvez utiliser Power Automate pour effectuer des opérations telles que l’ajout du contenu d’un e-mail à la table d’une feuille de calcul ou la création d’actions dans vos outils de gestion de projet en fonction des commentaires du classeur.

## <a name="get-started"></a>Prise en main

Si vous débutez avec Power Automate, nous vous recommandons de vous familiariser [avec Power Automate](/power-automate/getting-started). Vous y trouverez plus d’informations sur toutes les possibilités d’automatisation qui s’offrent à vous. Les documents ici se concentrent sur la façon dont les scripts Office fonctionnent avec Power Automate et sur la façon dont cela peut aider à améliorer votre expérience Excel.

### <a name="step-by-step-tutorials"></a>Didacticiels pas à pas

Il existe trois didacticiels pas à pas pour Power Automate et les scripts Office. Celles-ci montrent comment combiner les services d’automatisation et transmettre des données entre un classeur et un flux.

- [Appeler des scripts à partir d’un flux manuel Power Automate](../tutorials/excel-power-automate-manual.md)
- [Transmettre des données à des scripts dans un flux automatique Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement](../tutorials//excel-power-automate-returns.md)

### <a name="create-a-flow-from-excel"></a>Créer un flux à partir d’Excel

Vous pouvez commencer à utiliser Power Automate dans Excel avec différents modèles de flux. Sous l’onglet **Automatiser** , sélectionnez **Automatiser une tâche**.

:::image type="content" source="../images/automate-a-task-button.png" alt-text="Bouton « Automatiser une tâche » dans le ruban.":::

Cela ouvre un volet Office avec plusieurs options pour commencer à connecter vos scripts Office à des solutions automatisées plus volumineuses. Sélectionnez n’importe quelle option pour commencer. Votre flux est fourni avec le classeur actuel.

:::image type="content" source="../images/automate-a-task-choices.png" alt-text="Volet Office montrant les options de modèle de flux telles que « Planifier l’exécution d’un script Office dans Excel, puis envoyer un e-mail » et « Exécuter un script Office dans Excel lorsqu’une réponse Microsoft Forms est reçue ».":::

> [!TIP]
> Vous pouvez également commencer à créer un flux à partir du menu **Plus d’options (...)** sur un script individuel.

## <a name="excel-online-business-connector"></a>Connecteur Excel Online (Entreprise)

[Les connecteurs](/connectors/connectors) sont les ponts entre Power Automate et les applications. Le [connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness) permet à vos flux d’accéder aux classeurs Excel. L’action « Exécuter le script » vous permet d’appeler n’importe quel script Office accessible via le classeur sélectionné. Vous pouvez également donner à vos scripts des paramètres d’entrée afin que les données puissent être fournies par le flux, ou avoir vos informations de retour de script pour les étapes ultérieures du flux.

> [!IMPORTANT]
> L’action « Exécuter le script » donne aux personnes qui utilisent le connecteur Excel un accès significatif à votre classeur et à ses données. En outre, il existe des risques de sécurité avec les scripts qui effectuent des appels d’API externes, comme expliqué dans [les appels externes à partir de Power Automate](external-calls.md). Si votre administrateur est concerné par l’exposition de données hautement sensibles, il peut désactiver le connecteur Excel Online ou restreindre l’accès aux scripts Office via les [contrôles d’administrateur des scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

> [!IMPORTANT]
> Power Automate ne prend **pas** en charge les scripts stockés sur SharePoint pour l’instant.

## <a name="data-transfer-in-flows-for-scripts"></a>Transfert de données dans les flux pour les scripts

Power Automate vous permet de passer des éléments de données entre les étapes de votre flux. Les scripts peuvent être configurés pour accepter les types d’informations dont vous avez besoin et retourner tout ce que vous souhaitez dans votre classeur dans votre flux. L’entrée de votre script est spécifiée en ajoutant des paramètres à la `main` fonction (en plus de `workbook: ExcelScript.Workbook`). La sortie du script est déclarée en ajoutant un type de retour à `main`.

> [!NOTE]
> Lorsque vous créez un bloc « Exécuter un script » dans votre flux, les paramètres acceptés et les types retournés sont renseignés. Si vous modifiez les paramètres ou les types de retour de votre script, vous devez rétablir le bloc « Exécuter le script » de votre flux. Cela garantit que les données sont analysées correctement.

Les sections suivantes couvrent les détails de l’entrée et de la sortie des scripts utilisés dans Power Automate. Si vous souhaitez une approche pratique pour apprendre cette rubrique, essayez le didacticiel [Passer des données aux scripts dans un didacticiel de flux Power Automate exécuté automatiquement](../tutorials/excel-power-automate-trigger.md) ou explorez l’exemple de scénario [de rappels de tâches automatisés](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Paramètres : transmettre des données à un script

Toutes les entrées de script sont spécifiées en tant que paramètres supplémentaires pour la `main` fonction. Par exemple, si vous souhaitez qu’un script accepte un `string` nom qui représente un nom en tant qu’entrée, vous devez remplacer la `main` signature `function main(workbook: ExcelScript.Workbook, name: string)`par .

Lorsque vous configurez un flux dans Power Automate, vous pouvez spécifier une entrée de script en tant que valeurs statiques, [expressions](/power-automate/use-expressions-in-conditions) ou contenu dynamique. Pour plus d’informations sur le connecteur d’un service individuel, consultez la [documentation du connecteur Power Automate](/connectors/).

#### <a name="type-restrictions"></a>Restrictions de type

Lorsque vous ajoutez des paramètres d’entrée à la fonction d’un `main` script, tenez compte des allocations et restrictions suivantes. Elles s’appliquent également au type de retour du script.

1. Le premier paramètre doit être de type `ExcelScript.Workbook`. Son nom de paramètre n’a pas d’importance.

1. Les types `string`, `number`, `boolean`, `unknown`, `object`et `undefined` sont pris en charge.

1. Les tableaux (à la fois `[]` et `Array<T>` les styles) des types précédemment répertoriés sont pris en charge. Les tableaux imbriqués sont également pris en charge.

1. Les types Union sont autorisés s’il s’agit d’une union de littéraux appartenant à un seul type (par `"Left" | "Right"`exemple, pas `"Left", 5`). Les unions d’un type pris en charge avec un type non défini sont également prises en charge (par `string | undefined`exemple).

1. Les types d’objets sont autorisés s’ils contiennent des propriétés de type `string`, `number`, `boolean`des tableaux pris en charge ou d’autres objets pris en charge. L’exemple suivant montre des objets imbriqués qui sont pris en charge en tant que types de paramètres.

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

1. Leur interface ou définition de classe doit être définie dans le script. Un objet peut également être défini de façon anonyme inline, comme dans l’exemple suivant.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

#### <a name="optional-and-default-parameters"></a>Paramètres facultatifs et par défaut

1. Les paramètres facultatifs sont autorisés et sont indiqués avec le modificateur `?` facultatif (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)`).

1. Les valeurs de paramètre par défaut sont autorisées (par exemple `function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.

### <a name="return-data-from-a-script"></a>Retourner des données à partir d’un script

Les scripts peuvent retourner des données du classeur à utiliser comme contenu dynamique dans un flux Power Automate. Les [mêmes restrictions de type répertoriées précédemment](#type-restrictions) s’appliquent au type de retour. Pour renvoyer un objet, ajoutez la syntaxe de type de retour à la `main` fonction. Par exemple, si vous souhaitez retourner une `string` valeur à partir du script, votre `main` signature serait `function main(workbook: ExcelScript.Workbook): string`.

## <a name="example"></a>Exemple

La capture d’écran suivante montre un flux Power Automate qui est déclenché chaque fois qu’un problème [GitHub](https://github.com/) vous est attribué. Le flux exécute un script qui ajoute le problème à un tableau dans un classeur Excel. S’il y a au moins cinq problèmes dans cette table, le flux envoie un rappel par e-mail.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="Éditeur de flux Power Automate montrant l’exemple de flux.":::

La `main` fonction du script spécifie l’ID du problème et le titre du problème en tant que paramètres d’entrée, et le script retourne le nombre de lignes dans la table de problèmes.

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

- [Appeler des scripts à partir d’un flux manuel Power Automate](../tutorials/excel-power-automate-manual.md)
- [Transmettre des données à des scripts dans un flux automatique Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement](../tutorials/excel-power-automate-returns.md)
- [Informations de dépannage pour Power Automate avec les scripts Office](../testing/power-automate-troubleshooting.md)
- [Prise en main de Power Automate](/power-automate/getting-started)
- [Documentation de référence sur le connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness/)
