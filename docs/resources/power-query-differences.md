---
title: Quand utiliser des scripts Power Query ou Office
description: Scénarios les plus adaptés aux plateformes Power Query scripts Office scripts.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: e91077d635d66dde692c129bdd4b2f32657d5283
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585904"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Quand utiliser des scripts Power Query ou Office

[Power Query](https://powerquery.microsoft.com) scripts Office scripts sont tous deux des solutions d’automatisation puissantes pour Excel. Les deux solutions permettent Excel les utilisateurs de nettoyer et de transformer des données dans des workbooks. Un script Power Query ou Office unique peut être actualisé et réexécuté sur de nouvelles données pour produire des résultats cohérents, ce qui vous permet de gagner du temps et de travailler plus rapidement avec les informations qui en résultent.

Cet article fournit une vue d’ensemble générale du moment où vous pouvez privilégier une plateforme par rapport à l’autre. En règle générale, Power Query permet d’extrayer et de transformer des données à partir de sources de données externes et de scripts Office de grande taille, ce qui permet d’intégrer rapidement des solutions Excel et des Power Automate [externes](../develop/power-automate-integration.md).

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Sources de données et récupération de données importantes : Power Query

Nous vous recommandons de Power Query lorsque vous traitez des sources de données à partir de plateformes prises en charge.

Power Query des [connexions de données intégrées à](https://powerquery.microsoft.com/connectors/) des centaines de sources. Power Query est spécialement conçu pour la récupération, la transformation et les tâches de combinaison de données. Lorsque vous avez besoin de données provenant de l’une de ces sources, Power Query vous offre un moyen sans code d’apporter ces données dans Excel la forme dont vous avez besoin.

Ces Power Query sont conçues pour les jeux de données de grande taille. Ils n’ont pas les mêmes [limites](../testing/platform-limits.md) de transfert que Power Automate ou Excel sur le Web.

Office scripts offrent une solution légère pour les sources de données plus petites ou les sources de données non couvertes par les connecteurs Power Query de données. Cela [inclut l’utilisation `fetch` ou les API REST](../develop/external-calls.md) ou l’obtention d’informations à partir de sources de données ad hoc, telles qu’Teams [carte adaptative](../resources/scenarios/task-reminders.md).

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Mise en forme, visualisations et contrôle programmatique : Office scripts

Nous vous recommandons Office scripts lorsque vos besoins vont au-delà de l’importation et de la transformation des données.

Presque tout ce que vous pouvez faire manuellement via l’interface Excel’interface utilisateur peut être fait à l’Office scripts. Ils sont parfaits pour appliquer une mise en forme cohérente aux workbooks. Les scripts créent des graphiques, des tableaux croisés dynamiques, des formes, des images et d’autres visualisations de feuille de calcul. Les scripts vous donnent également un contrôle précis sur les positions, tailles, couleurs et autres attributs de ces visualisations.

L’inclusion de code TypeScript vous donne un haut degré de personnalisation. Une logique de contrôle programmatique telle `if...else` que des instructions rend votre script robuste. Cela vous permet d’apporter des opérations telles que lire des données de manière conditionnable sans vous appuyer sur des formules Excel complexes, ou d’analyser le workbook pour y voir des modifications inattendues avant de le modifier.

La mise en forme peut être appliquée Power Query modèles Excel [modèles.](https://templates.office.com/power-query-tutorial-tm11414620) Toutefois, les modèles sont mis à jour au niveau individuel ou de l’organisation, tandis que Office Scripts offrent un contrôle d’accès plus granulaire.

## <a name="power-automate-integrations"></a>Power Automate’intégrations

Office scripts offrent davantage d’options pour Power Automate’intégration. Les scripts sont adaptés à vos solutions. Vous définissez [l’entrée et la sortie du script](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts), afin qu’il fonctionne avec tout autre connecteur ou toute autre donnée dans le flux. La capture d’écran suivante montre un exemple Power Automate flux de données qui transmet les données d’une carte adaptative Teams à un script Office.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Capture d’écran  shows the Excel Online (Business) connector in the flow designer. Le connecteur utilise l’action de script Exécuter pour prendre une entrée à partir d Teams carte adaptative et la fournir à un script.":::

Power Query est utilisé dans le [connecteur SQL Server](https://powerquery.microsoft.com/flow/) Power Automate’une autre. Les [données Transform utilisant Power Query](/connectors/sql/#transform-data-using-power-query) action vous permet de créer une requête dans Power Automate. Bien qu’il s’agit d’un outil puissant à utiliser avec SQL Server, il limite les Power Query à cette source d’entrée, comme illustré dans la capture d’écran de flux suivante.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Capture d’écran  shows the SQL Server connector in the flow designer. Le connecteur utilise les données Transform à l’aide Power Query action.":::

## <a name="platform-dependencies"></a>Dépendances de plateforme

Office Scripts est actuellement disponible uniquement pour les Excel sur le Web. Power Query est actuellement disponible uniquement pour les Excel sur ordinateur de bureau. Les deux peuvent être utilisés par Power Automate, ce qui permet au flux de fonctionner avec Excel de travail stockés dans OneDrive.

## <a name="see-also"></a>Voir aussi

- [Power Query Portal](https://powerquery.microsoft.com/)
- [Power Query avec Excel](https://powerquery.microsoft.com/excel/)
- [Exécuter Office scripts avec Power Automate](../develop/power-automate-integration.md)
