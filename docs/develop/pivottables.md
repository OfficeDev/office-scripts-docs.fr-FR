---
title: Utiliser des tableaux croisés dynamiques dans des scripts Office
description: Découvrez le modèle objet pour les tableaux croisés dynamiques dans l’API JavaScript Office Scripts.
ms.date: 04/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: 579f94140214674912c9610e707123924e4aef18
ms.sourcegitcommit: 4e3d3aa25fe4e604b806fbe72310b7a84ee72624
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/27/2022
ms.locfileid: "65077064"
---
# <a name="work-with-pivottables-in-office-scripts"></a>Utiliser des tableaux croisés dynamiques dans des scripts Office

Les tableaux croisés dynamiques vous permettent d’analyser rapidement de grandes collections de données. Avec leur puissance vient la complexité. Les API Office Scripts vous permettent de personnaliser un tableau croisé dynamique en fonction de vos besoins, mais l’étendue de l’ensemble d’API rend la prise en main difficile. Cet article montre comment effectuer des tâches courantes de tableau croisé dynamique et explique les classes et méthodes importantes.

> [!NOTE]
> Pour mieux comprendre le contexte des termes utilisés par les API, lisez d’abord la documentation du tableau croisé dynamique de Excel. Commencez par [créer un tableau croisé dynamique pour analyser les données de feuille de calcul](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576).

## <a name="object-model"></a>Modèle d’objet

:::image type="content" source="../images/pivottable-object-model.png" alt-text="Image simplifiée des classes, méthodes et propriétés utilisées lors de l’utilisation de tableaux croisés dynamiques.":::

Le [tableau croisé dynamique](/javascript/api/office-scripts/excelscript/excelscript.pivottable) est l’objet central des tableaux croisés dynamiques dans l’API Office Scripts.

- [L’objet Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) a une collection de tous les [tableaux croisés dynamiques](/javascript/api/office-scripts/excelscript/excelscript.pivottable). Chaque [feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contient également une collection de tableaux croisés dynamiques locale à cette feuille.
- Un [tableau croisé dynamique](/javascript/api/office-scripts/excelscript/excelscript.pivottable) contient des [tableaux croisés dynamiques](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy). Une hiérarchie peut être considérée comme une colonne dans une table.
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) peut être ajouté en tant que lignes ou colonnes ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)), données ([DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)) ou filtres ([FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)).
- Chaque [pivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) contient exactement un [champ de tableau croisé dynamique](/javascript/api/office-scripts/excelscript/excelscript.pivotfield). Les structures de tableau croisé dynamique en dehors de Excel peuvent contenir plusieurs champs par hiérarchie. Cette conception existe donc pour prendre en charge les options futures. Pour Office scripts, les champs et les hiérarchies correspondent aux mêmes informations.
- Un [champ de tableau](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) croisé dynamique contient plusieurs éléments de tableau [croisé dynamique](/javascript/api/office-scripts/excelscript/excelscript.pivotitem). Chaque pivotItem est une valeur unique dans le champ. Considérez chaque élément comme une valeur dans la colonne de table. Les éléments peuvent également être des valeurs agrégées, telles que des sommes, si le champ est utilisé pour les données.
- [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) définit la façon dont les champs de tableau croisé dynamique et [les](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) éléments [dynamiques sont affichés](/javascript/api/office-scripts/excelscript/excelscript.pivotitem).
- [Les filtres de tableau croisé dynamique filtrent les](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) données du [tableau croisé dynamique](/javascript/api/office-scripts/excelscript/excelscript.pivottable) à l’aide de différents critères.

Examinez le fonctionnement de ces relations dans la pratique. Les données suivantes décrivent les ventes de fruits provenant de différentes fermes. Il s’agit de la base de tous les exemples de cet article. Utilisez <a href="pivottable-sample.xlsx">pivottable-sample.xlsx</a> pour suivre.

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="Collection de ventes de fruits de différents types provenant de différentes fermes.":::

## <a name="create-a-pivottable-with-fields"></a>Créer un tableau croisé dynamique avec des champs

Les tableaux croisés dynamiques sont créés avec des références à des données existantes. Les plages et les tables peuvent être la source d’un tableau croisé dynamique. Ils ont également besoin d’un emplacement pour exister dans le classeur. Étant donné que la taille d’un tableau croisé dynamique est dynamique, seul le coin supérieur gauche de la plage de destination est spécifié.

L’extrait de code suivant crée un tableau croisé dynamique basé sur une plage de données. Le tableau croisé dynamique n’a aucune hiérarchie. Les données ne sont donc pas encore regroupées d’une quelconque manière.

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="Tableau croisé dynamique nommé « Farm Pivot » sans hiérarchie.":::

### <a name="hierarchies-and-fields"></a>Hiérarchies et champs

Les tableaux croisés dynamiques sont organisés par le biais de hiérarchies. Ces hiérarchies sont utilisées pour pivoter les données lorsqu’elles sont ajoutées en tant que type spécifique de hiérarchie. Il existe quatre types de hiérarchies.

- **Ligne** : affiche les éléments dans des lignes horizontales.
- **Colonne** : affiche les éléments dans des colonnes verticales.
- **Données** : affiche des agrégats de valeurs en fonction des lignes et des colonnes.
- **Filtre** : ajouter ou supprimer des éléments du tableau croisé dynamique.

Un tableau croisé dynamique peut avoir autant ou moins de champs affectés à ces hiérarchies spécifiques. Un tableau croisé dynamique a besoin d’au moins une hiérarchie de données pour afficher les données numériques résumées et d’au moins une ligne ou colonne sur laquelle pivoter ce résumé. L’extrait de code suivant ajoute deux hiérarchies de lignes et deux hiérarchies de données.

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="Tableau croisé dynamique montrant le total des ventes de différents fruits en fonction de la ferme d’où ils proviennent.":::

## <a name="layout-ranges"></a>Plages de disposition

Chaque partie du tableau croisé dynamique est mappée à une plage. Cela permet à votre script d’obtenir des données à partir du tableau croisé dynamique pour une utilisation ultérieure dans le script ou pour être retourné dans un [flux de Power Automate](power-automate-integration.md). Ces plages sont accessibles via l’objet [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) acquis à partir de `PivotTable.getLayout()`. Le diagramme suivant montre les plages retournées par les méthodes dans `PivotLayout`.

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="Diagramme montrant les sections d’un tableau croisé dynamique retournées par les fonctions get range de la disposition.":::

## <a name="filters-and-slicers"></a>Filtres et segments

Il existe trois façons de filtrer un tableau croisé dynamique.

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### <a name="filterpivothierarchies"></a>FilterPivotHierarchies

`FilterPivotHierarchies` ajouter une hiérarchie supplémentaire pour filtrer chaque ligne de données. Toute ligne avec un élément filtré est exclue du tableau croisé dynamique et de ses résumés. Étant donné que ces filtres sont basés sur des éléments, ils fonctionnent uniquement sur des valeurs discrètes. Si « Classification » est une hiérarchie de filtres dans notre exemple, les utilisateurs peuvent sélectionner les valeurs « Organique » et « Conventionnel » pour le filtre. De même, si « Crates Sold Wholesale » est sélectionné, les options de filtre sont les nombres individuels, tels que 120 et 150, au lieu des plages numériques.

`FilterPivotHierarchies` sont créés avec toutes les valeurs sélectionnées. Cela signifie que rien n’est filtré tant que l’utilisateur n’interagit pas manuellement avec le contrôle de filtre ou qu’un `PivotManualFilter` n’est pas défini sur le champ appartenant au `FilterPivotHierarchy`.

L’extrait de code suivant ajoute « Classification » en tant que hiérarchie de filtre.

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="Contrôle de filtre qui utilise « Classification » pour un tableau croisé dynamique.":::

### <a name="pivotfilters"></a>PivotFilters

L’objet `PivotFilters` est une collection de filtres appliqués à un seul champ. Étant donné que chaque hiérarchie a exactement un champ, vous devez toujours utiliser le premier champ dans `PivotHierarchy.getFields()` lors de l’application de filtres. Il existe quatre types de filtres.

- **Filtre de date** : filtrage basé sur les dates du calendrier.
- **Filtre d’étiquette** : filtrage de comparaison de texte.
- **Filtre manuel** : filtrage d’entrée personnalisé.
- **Filtre de valeur** : filtrage de comparaison de nombres. Cela compare les éléments de la hiérarchie associée aux valeurs d’une hiérarchie de données spécifiée.

En règle générale, un seul des quatre types de filtres est créé et appliqué au champ. Si le script tente d’utiliser des filtres incompatibles, une erreur est générée avec le texte « L’argument n’est pas valide ou manquant ou a un format incorrect ».

L’extrait de code suivant ajoute deux filtres. Le premier est un filtre manuel qui sélectionne des éléments dans une hiérarchie de filtres « Classification » existante. Le deuxième filtre supprime toutes les fermes qui ont moins de 300 « Crates Sold Wholesale ». Notez que cela filtre la « somme » de ces batteries de serveurs, et non les lignes individuelles des données d’origine.

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="Tableau croisé dynamique après l’application du filtre de valeurs et du filtre manuel.":::

### <a name="slicers"></a>Slicers

[Les segments filtrent les](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) données dans un tableau croisé dynamique (ou une table standard). Il s’agit d’objets pouvant être déplacés dans la feuille de calcul qui permettent de filtrer rapidement les sélections. Un segment fonctionne de la même façon que le filtre manuel et `PivotFilterHierarchy`. Les éléments du `PivotField` tableau croisé dynamique sont activés pour les inclure ou les exclure du tableau croisé dynamique.

L’extrait de code suivant ajoute un segment pour le champ « Type ». Il définit les éléments sélectionnés sur « Lemon » et « Lime », puis déplace le segment de 400 pixels vers la gauche.

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="Segment filtrant les données sur un tableau croisé dynamique.":::

## <a name="see-also"></a>Voir aussi

- [Principes de base pour la rédaction de scripts Office en Excel sur le web](scripting-fundamentals.md)
- [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview)
