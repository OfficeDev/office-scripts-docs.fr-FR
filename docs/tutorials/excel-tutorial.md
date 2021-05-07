---
title: Enregistrer, modifier, créer des scripts Office dans Excel pour le web
description: Didacticiel sur les notions de base des scripts Office, comprenant l’enregistrement de scripts avec l’enregistreur d’actions et l’écriture de données dans un classeur.
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: 9f1b2e29d60ec0e370bdb29fde0f04be831a222b
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232864"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a>Enregistrer, modifier, créer des scripts Office dans Excel pour le web

Ce didacticiel vous présente les notions de base de l’enregistrement, de la modification et de la rédaction d’un script Office pour Excel sur le web. Vous allez enregistrer un script mettant en forme une feuille de calcul d’enregistrement des ventes. Vous allez ensuite modifier le script enregistré pour appliquer une mise en forme supplémentaire, créer un tableau, puis trier ce tableau. Ce modèle de type « enregistrement suivi d’une modification » constitue un outil important pour vous permettre de savoir à quoi ressemblent vos actions Excel sous forme de code.

## <a name="prerequisites"></a>Conditions préalables

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> Ce didacticiel est destiné aux utilisateurs ayant des connaissances de niveau débutant à intermédiaire en JavaScript ou TypeScript. Si vous découvrez JavaScript, nous vous conseillons de commencer par consulter le [didacticiel Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Si vous souhaitez en savoir plus sur l’environnement de script, veuillez consulter la rubrique [Environnement de l’éditeur de code Scripts Office](../overview/code-editor-environment.md).

## <a name="add-data-and-record-a-basic-script"></a>Ajouter des données et enregistrer un script simple

Tout d’abord, il nous faut des données et un petit script de base.

1. Créez un nouveau classeur dans Excel pour le Web.
2. Copiez les données de ventes de fruits suivantes et collez-les dans la feuille de calcul en commençant à la cellule **A1**.

    |Fruits |2018 |2019 |
    |:---|:---|:---|
    |Oranges |1000 |1200 |
    |Citrons |800 |900 |
    |Citrons verts |600 |500 |
    |Pamplemousses |900 |700 |

3. Ouvrez l’onglet **Automatiser**. Si vous ne voyez pas l’onglet **Automatiser**, vérifiez dans la section dépassement du ruban en appuyant sur la flèche déroulante vers le bas.
4. Appuyez sur le bouton **Actions d’enregistrement**.
5. Sélectionnez les cellules **A2:C2** (la ligne « Oranges ») et choisissez orange comme couleur de remplissage.
6. Appuyez sur le bouton **Arrêter** pour arrêter l’enregistrement.
7. Renseignez le champ **Nom du script** avec un nom explicite.
8. *Facultatif :* renseignez le champ **Description** avec une description significative. Celle-ci permet d’offrir un contexte sur l’usage du script. Pour ce didacticiel, vous pouvez utiliser « Assigner un code couleur aux lignes d’un tableau ».

   > [!TIP]
   > Vous pouvez modifier la description d’un script ultérieurement à partir du volet **Détails du script** qui se trouve sous le menu **...** de l’Éditeur de code.

9. Sauvegardez le script en cliquant sur le bouton **Enregistrer**.

    Voici ce à quoi votre feuille de calcul doit ressembler (les couleurs peuvent être différentes) :

    :::image type="content" source="../images/tutorial-1.png" alt-text="Feuille de calcul affichant une ligne de données de ventes de fruits avec les « Oranges » mises en évidence par la couleur orange.":::

## <a name="edit-an-existing-script"></a>Modifier un script existant

Le script précédent a coloré la ligne « Oranges » en orange. Nous allons ajouter une ligne jaune pour « Citrons ».

1. Depuis le volet **Détails** à présent ouvert, appuyez sur le bouton **Modifier**.
2. Un code similaire à celui-ci doit apparaître :

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let selectedSheet = workbook.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
    }
    ```

    Ce code extrait la feuille de calcul actuelle du classeur. Il définit ensuite la couleur de remplissage de la plage **A2:C2**.

    Les plages jouent un rôle fondamental dans les scripts Office d’Excel pour le web. Une plage est un bloc de cellules contiguës de forme rectangulaire qui contient des valeurs, des formules ou des formats. Les plages constituent la structure de base faite de cellules par laquelle vous effectuerez des tâches de script.

3. Ajoutez la ligne suivante à la fin du script (entre l’emplacement où le `color` se trouve et le `}` de clôture) :

    ```TypeScript
    selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
    ```

4. Testez le script en appuyant sur **Exécuter**. Voici ce à quoi votre feuille de calcul doit maintenant ressembler :

    :::image type="content" source="../images/tutorial-2.png" alt-text="Feuille de calcul affichant la ligne des données de ventes de fruits avec la ligne « Oranges » mise en évidence par la couleur orange et la ligne « Citrons » par la couleur jaune.":::

## <a name="create-a-table"></a>Créer un tableau

Nous allons convertir les données de ventes de fruits en tableau. Nous allons utiliser notre script pour l’ensemble du processus.

1. Ajoutez la ligne suivante à la fin du script (avant le `}` de clôture) :

    ```TypeScript
    let table = selectedSheet.addTable("A1:C5", true);
    ```

2. Cet appel renvoie un objet `Table`. Nous allons utiliser ce tableau pour trier les données. Nous allons trier les données en ordre croissant en fonction des valeurs de la colonne « Fruits ». Ajoutez la ligne suivante après la création du tableau :

    ```TypeScript
    table.getSort().apply([{ key: 0, ascending: true }]);
    ```

    Voici ce à quoi doit ressembler votre script :

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Set fill color to FFC000 for range Sheet1!A2:C2
        let selectedSheet = workbook.getActiveWorksheet();
        selectedSheet.getRange("A2:C2").getFormat().getFill().setColor("FFC000");
        selectedSheet.getRange("A3:C3").getFormat().getFill().setColor("yellow");
        let table = selectedSheet.addTable("A1:C5", true);
        table.getSort().apply([{ key: 0, ascending: true }]);
    }
    ```

    Les tableaux comportent un objet `TableSort`, accessible via la méthode `Table.getSort`. Vous pouvez appliquer des critères de tri à cet objet. La méthode `apply` prend un tableau d’objets `SortField`. Dans notre cas, ne disposant que d’un seul critère de tri, nous utiliserons un seul `SortField`. `key: 0` définit la colonne ayant les valeurs de définition de tri sur « 0 » (la première colonne du tableau, **A** dans notre cas). `ascending: true` trie les données dans un ordre croissant (et non dans un ordre décroissant).

3. Exécutez le script. Un tableau come ceci devrait s’afficher :

    :::image type="content" source="../images/tutorial-3.png" alt-text="Feuille de calcul affichant la table de ventes des fruits triées.":::

    > [!NOTE]
    > Si vous réexécutez le script, un message d’erreur s’affiche. En effet, vous ne pouvez pas créer un tableau au-dessus d’un autre. Toutefois, vous pouvez exécuter le script sur une autre feuille de calcul ou un autre classeur.

### <a name="re-run-the-script"></a>Réexécutez le script.

1. Créer une nouvelle feuille de calcul dans le classeur actif.
2. Copiez les données des fruits du début de ce didacticiel et collez-les dans la nouvelle feuille de calcul, en commençant à la cellule **A1**.
3. Exécutez le script.

## <a name="next-steps"></a>Étapes suivantes

Complétez le didacticiel [Lire les données d’un classeur avec les scripts Office d’Excel pour le web](excel-read-tutorial.md). Il vous apprend comment lire des données à partir d’un classeur à l’aide d’un script Office.
