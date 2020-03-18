---
title: Lire les données d’un classeur avec les scripts Office d’Excel pour le web
description: Didacticiel des scripts Office sur la lecture de données à partir de classeurs et l’évaluation de ces données dans le script.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 42ed0fe5843a78692f9660b873211e3668702164
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700181"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>Lire les données d’un classeur avec les scripts Office d’Excel pour le web

Ce didacticiel vous apprend comment lire des données à partir d’un classeur à l’aide d’un script Office pour Excel pour le web. Vous pourrez ensuite modifier les données que vous avez lues et les replacer dans le classeur.

> [!TIP]
> Si vous débutez avec les scripts Office, nous vous recommandons de commencer par le didacticiel [Enregistrer, modifier, créer des scripts Office dans Excel pour le web](excel-tutorial.md).

## <a name="prerequisites"></a>Conditions préalables

[!INCLUDE [Preview note](../includes/preview-note.md)]

Avant de commencer ce didacticiel, vous devez disposer d’un accès aux scripts Office, ce qui nécessite ce qui suit :

- [Excel pour le web](https://www.office.com/launch/excel).
- Demandez à votre administrateur d’[activer les scripts Office pour votre organisation](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), ce qui ajoute l’onglet **Automatiser** au ruban.

> [!IMPORTANT]
> Ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript. Si vous débutez avec JavaScript, nous vous conseillons de consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Rendez-vous sur [Scripts Office dans Excel pour le web](../overview/excel.md) pour en savoir plus sur l’environnement de script.

## <a name="read-a-cell"></a>Lire une cellule

Les scripts créés avec l’enregistreur d’actions peuvent uniquement écrire des informations dans le classeur. L’éditeur de code vous permet de modifier et de créer des scripts qui peuvent également lire les données d’un classeur.

Nous allons créer un script qui lit les données et agit en fonction de ce qui a été lu. Nous allons utiliser un exemple de relevé bancaire. Il s’agit d’un relevé combiné de compte courant et de crédit. Malheureusement, les changements de soldes sont rapportés différemment. Le relevé de compte courant donne les revenus comme crédit positif et les dépenses comme débit négatif. Le relevé de crédit fait l’inverse.

Dans le reste du didacticiel, nous allons normaliser ces données à l’aide d’un script. Pour commencer, voyons comment lire des données à partir du classeur.

1. Créez une nouvelle feuille de calcul dans le classeur courant, vous l’utiliserez pour le reste du didacticiel.
2. Copiez les données suivantes et collez-les dans la feuille de calcul en commençant à la cellule **A1**.

    |Date |Compte |Description |Débit |Crédit |
    |:--|:--|:--|:--|:--|
    |10/10/2019 |Compte courant |Coho Vineyard |−20,05 | |
    |11/10/2019 |Crédit |The Phone Company |99,95 | |
    |13/10/2019 |Crédit |Coho Vineyard |154,43 | |
    |15/10/2019 |Compte courant |Versement externe | |1000 |
    |20/10/2019 |Crédit |Coho Vineyard − Remboursement | |−35,45 |
    |25/10/2019 |Compte courant |Best For You Organics Company | −85,64 | |
    |01/11/2019 |Compte courant |Versement externe | |1000 |

3. Ouvrez l’**éditeur de code** puis sélectionnez **Nouveau script**.
4. Nous allons réarranger la mise en forme. Il s’agit d’un document financier, nous allons donc modifier la mise en forme des nombres dans les colonnes **Débit** et **Crédit** pour afficher les valeurs sous forme de montants en dollars. Ajustons également la largeur des colonnes aux données.

    Remplacez le contenu du script par le code suivant :

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

5. Nous allons maintenant lire une valeur depuis l’une des colonnes de montants. Ajoutez le code suivant à la fin du script :

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    range.load("values");
    await context.sync();
  
    // Print the value of D2.
    console.log(range.values);
    ```

    Remarquez les appels de `load` et de `sync`. Pour plus de détails sur ces méthodes, voir [Principes de base des scripts Office dans Excel pour le web](../develop/scripting-fundamentals.md#sync-and-load). Pour l’instant, sachez seulement que vous devez demander la lecture des données puis synchroniser votre script avec le classeur pour lire les données.

6. Exécutez le script.
7. Ouvrez la console. Accédez au menu **Ellipses**, puis appuyez sur **Journaux...**.
8. Dans la console, `[Array[1]]` doit s’afficher. Ce n’est pas un nombre, car les plages sont des tableaux de données à deux dimensions. Cette plage à deux dimensions est directement journalisée dans la console. Heureusement, l’éditeur de code vous permet de voir le contenu du tableau.
9. Lorsqu’un tableau à deux dimensions est journalisé sur la console, il regroupe les valeurs de colonne sous chaque ligne. Développez le journal du tableau en appuyant sur le triangle bleu.
10. Développez le deuxième niveau du tableau en appuyant sur le triangle bleu nouvellement affiché. Voici ce que vous devez voir :

    ![Journal de la console affichant la sortie « −20,05 », imbriquée sous deux tableaux.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>Modifier la valeur d’une cellule.

Maintenant que nous avons vu comment lire des données, nous allons les utiliser pour modifier le classeur. Nous allons rendre la valeur de la cellule **D2** positive avec la fonction `Math.abs`. L’objet [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contient de nombreuses fonctions auxquelles vos scripts ont accès. Pour plus d’informations sur `Math` et les autres objets intégrés, voir [Utilisation d’objets JavaScript intégrés dans les scripts Office](../develop/javascript-objects.md).

1. Ajoutez le code suivant à la fin du script :

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.values[0][0]);
    range.values = [[positiveValue]];
    ```

2. La valeur de la cellule **D2** doit maintenant être positive.

## <a name="modify-the-values-of-a-column"></a>Modifier les valeurs d’une colonne

Maintenant que nous avons vu comment lire et écrire dans une seule cellule, configurons le script de façon à ce qu’il travaille sur l’ensemble des cellules des colonnes **Débit** et **Crédit**.

1. Supprimez le code qui affecte une seule cellule (le code de valeur absolue précédent), de sorte que votre script se présente désormais comme suit :

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the current worksheet.
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();

      // Format the range to display numerical dollar amounts.
      selectedSheet.getRange("D2:E8").numberFormat = [["$#,##0.00"]];

      // Fit the width of all the used columns to the data.
      selectedSheet.getUsedRange().format.autofitColumns();
    }
    ```

2. Ajoutez une boucle pour produire une itération dans des lignes des deux dernières colonnes. Le script remplace la valeur de chaque cellule en la valeur absolue de cette valeur.

    Notez que l’indexation du tableau qui définit les emplacements des cellules est basée sur zéro. Par conséquent, la cellule **A1** est `range[0][0]`.

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    range.load("rowCount,values");
    await context.sync();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.rowCount; i++) {
      // The column at index 3 is column "4" in the worksheet.
      if (range.values[i][3] != 0) {
        let positiveValue = Math.abs(range.values[i][3]);
        selectedSheet.getCell(i, 3).values = [[positiveValue]];
      }

      // The column at index 4 is column "5" in the worksheet.
      if (range.values[i][4] != 0) {
        let positiveValue = Math.abs(range.values[i][4]);
        selectedSheet.getCell(i, 4).values = [[positiveValue]];
      }
    }
    ```

    Cette partie du script effectue plusieurs tâches importantes. Premièrement, elle charge les valeurs et le nombre de lignes de la plage utilisée. Nous pouvons ainsi examiner les valeurs et déterminer quand arrêter. Deuxièmement, elle produit une itération dans la plage utilisée, en vérifiant chaque cellule des colonnes **Débit** et **Crédit**. Enfin, si la valeur dans la cellule n’est pas 0, elle est remplacée par sa valeur absolue. Nous évitons les zéros pour pouvoir laisser les cellules vides telles qu’elles sont.

3. Exécutez le script.

    Voici ce à quoi doit maintenant ressembler le relevé bancaire :

    ![Le relevé bancaire sous la forme d’un tableau mis en forme avec uniquement des valeurs positives.](../images/tutorial-5.png)

## <a name="next-steps"></a>Étapes suivantes

Ouvrez l’éditeur de code et testez quelques-uns de nos [Exemples de scripts pour Scripts Office dans Excel pour le web](../resources/excel-samples.md). Vous pouvez également consulter [Principes de base des scripts Office dans Excel pour le web](../develop/scripting-fundamentals.md) pour en savoir plus sur la création de scripts Office.
