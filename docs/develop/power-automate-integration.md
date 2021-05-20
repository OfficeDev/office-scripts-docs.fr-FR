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
# <a name="run-office-scripts-with-power-automate"></a>Exécutez Office scripts avec Power Automate

[Power Automate vous](https://flow.microsoft.com) permet d’ajouter Office scripts à un flux de travail automatisé plus grand. Vous pouvez utiliser Power Automate des choses comme ajouter le contenu d’un e-mail à la table d’une feuille de travail ou créer des actions dans vos outils de gestion de projet s’appuyant sur des commentaires de manuel.

## <a name="get-started"></a>Prise en main

Si vous êtes nouveau à Power Automate, nous vous recommandons de [visiter Démarrer avec Power Automate](/power-automate/getting-started). Là, vous pouvez en apprendre davantage sur toutes les possibilités d’automatisation à votre disposition. Les documents ici se concentrent sur la façon dont Office scripts fonctionnent avec Power Automate et comment cela peut aider à améliorer votre expérience Excel vie.

Pour commencer à combiner Power Automate scripts Office, suivez le tutoriel [Commencez à utiliser des scripts avec Power Automate](../tutorials/excel-power-automate-manual.md). Cela vous apprendra à créer un flux qui appelle un script simple. Une fois que vous avez terminé ce tutoriel et [les données Pass aux scripts dans un tutoriel de flux de Power Automate exécuté](../tutorials/excel-power-automate-trigger.md) automatiquement, revenez ici pour plus d’informations détaillées sur la connexion des scripts Office aux flux Power Automate utilisateurs.

## <a name="excel-online-business-connector"></a>Excel Connecteur en ligne (Business)

[Les connecteurs](/connectors/connectors) sont les ponts entre Power Automate et les applications. Le [Excel connecteur en ligne (Business)](/connectors/excelonlinebusiness) donne à vos flux accès à Excel manuels. L’action « Exécuter le script » vous permet d’appeler n’importe Office script accessible via le cahier de travail sélectionné. Vous pouvez également donner à vos scripts des paramètres d’entrée afin que les données puissent être fournies par le flux, ou avoir vos informations de retour de script pour les étapes ultérieures du flux.

> [!IMPORTANT]
> L’action « Exécuter script » donne aux personnes qui utilisent le connecteur Excel un accès significatif à votre cahier de travail et à ses données. En outre, il ya des risques de sécurité avec les scripts qui font des appels API externes, comme expliqué [dans les appels externes de Power Automate](external-calls.md). Si votre administrateur est préoccupé par l’exposition de données hautement sensibles, ils peuvent soit désactiver le connecteur Excel Online ou restreindre l’accès aux scripts Office par le [biais des contrôles de l’administrateur scripts Office](/microsoft-365/admin/manage/manage-office-scripts-settings).

## <a name="data-transfer-in-flows-for-scripts"></a>Transfert de données dans les flux de scripts

Power Automate vous permet de transmettre des éléments de données entre les étapes de votre flux. Les scripts peuvent être configurés pour accepter tous les types d’informations dont vous avez besoin et renvoyer tout ce qui vient de votre cahier de travail que vous souhaitez dans votre flux. L’entrée de votre script est spécifiée en ajoutant des paramètres à `main` la fonction (en plus de `workbook: ExcelScript.Workbook` ). La sortie du script est déclarée en ajoutant un type de retour à `main` .

> [!NOTE]
> Lorsque vous créez un bloc « Script d’exécuter » dans votre flux, les paramètres acceptés et les types retournés sont remplis. Si vous modifiez les paramètres ou retournez les types de votre script, vous devrez refaire le bloc « Exécuter le script » de votre flux. Cela garantit que les données sont correctement analyses.

Les sections suivantes couvrent les détails de l’entrée et de la sortie des scripts utilisés dans Power Automate. Si vous souhaitez une approche pratique pour l’apprentissage de ce sujet, essayez les données Pass aux scripts dans un [didacticiel de flux de Power Automate exécuté](../tutorials/excel-power-automate-trigger.md) automatiquement ou explorez le scénario d’exemple de [rappels de tâches automatisés.](../resources/scenarios/task-reminders.md)

### <a name="main-parameters-pass-data-to-a-script"></a>`main` Paramètres : Transmettre des données à un script

Toutes les entrées de script sont spécifiées comme paramètres supplémentaires pour la `main` fonction. Par exemple, si vous vouliez qu’un script accepte un `string` nom qui représente un nom comme entrée, vous changeriez la signature en `main` `function main(workbook: ExcelScript.Workbook, name: string)` .

Lorsque vous configurez un flux dans les Power Automate, vous pouvez spécifier l’entrée du script sous forme de valeurs [statiques, d’expressions](/power-automate/use-expressions-in-conditions)ou de contenu dynamique. Les détails sur le connecteur d’un service individuel peuvent être trouvés [dans la documentation Power Automate Connecteur.](/connectors/)

Lorsque vous ajoutez des paramètres d’entrée à la fonction `main` d’un script, considérez les allocations et restrictions suivantes.

1. Le premier paramètre doit être de type `ExcelScript.Workbook` . Son nom de paramètre n’a pas d’importance.

2. Chaque paramètre doit avoir un type (tel `string` ou `number` ).

3. Les types de `string` base , , , , , et sont pris en `number` `boolean` `unknown` `object` `undefined` charge.

4. Les tableaux des types de base précédemment répertoriés sont pris en charge.

5. Les tableaux imbriqués sont pris en charge sous forme de paramètres (mais pas en tant que types de retour).

6. Les types d’union sont autorisés s’il s’agit d’une union de littérales appartenant à un seul type `"Left" | "Right"` (comme). Les unions d’un type soutenu avec des non définis sont également soutenues (telles que `string | undefined` ).

7. Les types d’objets sont autorisés s’ils contiennent des propriétés de `string` `number` type, `boolean` , des tableaux pris en charge, ou d’autres objets pris en charge. L’exemple suivant montre les objets imbriqués qui sont pris en charge sous forme de types de paramètres :

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

8. Les objets doivent avoir leur interface ou définition de classe définie dans le script. Un objet peut également être défini anonymement en ligne, comme dans l’exemple suivant :

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Les paramètres optionnels sont autorisés et peuvent être indiqués comme tels en utilisant le modificateur `?` optionnel (par exemple, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Les valeurs de paramètres par défaut sont autorisées (par exemple `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .

### <a name="return-data-from-a-script"></a>Renvoyer les données d’un script

Les scripts peuvent renvoyer des données du cahier de travail pour les utiliser comme contenu dynamique dans un flux Power Automate fluide. Comme pour les paramètres d’entrée, Power Automate impose certaines restrictions sur le type de retour.

1. Les types de base `string` , , , et sont pris en `number` `boolean` `void` `undefined` charge.

2. Les types d’union utilisés comme types de retour suivent les mêmes restrictions qu’ils le font lorsqu’ils sont utilisés comme paramètres de script.

3. Les types de tableaux sont autorisés s’ils sont de `string` `number` type, ou `boolean` . Ils sont également autorisés si le type est un syndicat soutenu ou soutenu type littéral.

4. Les types d’objets utilisés comme types de retour suivent les mêmes restrictions qu’ils le font lorsqu’ils sont utilisés comme paramètres de script.

5. La dactylographie implicite est prise en charge, bien qu’elle doit suivre les mêmes règles qu’un type défini.

## <a name="example"></a>Exemple

La capture d’écran suivante Power Automate un flux de flux qui est déclenché chaque [fois qu’GitHub](https://github.com/) problème est attribué à vous. Le flux exécute un script qui ajoute le problème à une table dans un Excel de travail. S’il y a cinq problèmes ou plus dans ce tableau, le flux envoie un rappel par courriel.

:::image type="content" source="../images/power-automate-parameter-return-sample.png" alt-text="L’éditeur Power Automate débit de flux montrant le flux d’exemple":::

La `main` fonction du script spécifie l’ID d’émission et le titre d’émission en tant que paramètres d’entrée, et le script renvoie le nombre de lignes dans le tableau d’émission.

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

- [Exécutez Office scripts en Excel sur le Web avec Power Automate](../tutorials/excel-power-automate-manual.md)
- [Transmettre des données à des scripts dans un flux automatique Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Renvoyer les données d’un script vers un flux Power Automate exécuté automatiquement](../tutorials/excel-power-automate-returns.md)
- [Informations de dépannage pour Power Automate avec Office scripts](../testing/power-automate-troubleshooting.md)
- [Prise en main de Power Automate](/power-automate/getting-started)
- [Excel Documentation de référence connecteur en ligne (Business)](/connectors/excelonlinebusiness/)
