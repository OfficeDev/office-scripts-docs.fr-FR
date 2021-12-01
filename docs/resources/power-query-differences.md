---
title: Quand utiliser Power Query ou Office Scripts
description: Scénarios les plus adaptés aux plateformes Power Query et Office Scripts.
ms.date: 11/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1812b508b2cde4d304ecf228adfdd8f68de9808a
ms.sourcegitcommit: 383880e0dc0d09b8f76884675531e462a292d747
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/01/2021
ms.locfileid: "61245612"
---
# <a name="when-to-use-power-query-or-office-scripts"></a>Quand utiliser Power Query ou Office Scripts

[Power Query et](https://powerquery.microsoft.com) Office Scripts sont tous deux des solutions d’automatisation puissantes pour Excel. Les deux solutions permettent Excel les utilisateurs de nettoyer et de transformer des données dans des workbooks. Un script Power Query ou Office unique peut être actualisé et réexécuté sur de nouvelles données pour produire des résultats cohérents, ce qui vous permet de gagner du temps et de travailler plus rapidement avec les informations qui en résultent.

Cet article fournit une vue d’ensemble générale du moment où vous pouvez privilégier une plateforme par rapport à l’autre. En règle générale, Power Query est bon pour tirer et transformer des données à partir de sources de données externes de grande taille et des scripts Office sont bons pour des intégrations rapides, [centrées](../develop/power-automate-integration.md)sur Excel et Power Automate.

## <a name="large-data-sources-and-data-retrieval-power-query"></a>Sources de données et récupération de données importantes : Power Query

Nous vous recommandons Power Query lorsque vous traitez des sources de données à partir de plateformes prises en charge.

Power Query dispose [de connexions de données intégrées à](https://powerquery.microsoft.com/connectors/) des centaines de sources. Power Query est spécialement conçu pour la récupération, la transformation et les tâches de combinaison de données. Lorsque vous avez besoin de données provenant de l’une de ces sources, Power Query vous offre un moyen sans code d’apporter ces données dans Excel la forme dont vous avez besoin.

Ces connexions Power Query sont conçues pour les jeux de données de grande taille. Ils n’ont pas les mêmes [limites de](../testing/platform-limits.md) transfert que Power Automate ou Excel sur le Web.

Office scripts offrent une solution légère pour les sources de données plus petites ou les sources de données non couvertes par les connecteurs Power Query. Cela [inclut `fetch` l’utilisation](../develop/external-calls.md) ou les API REST ou l’obtention d’informations à partir de sources de données ad hoc, telles qu’Teams [carte adaptative .](../resources/scenarios/task-reminders.md)

## <a name="formatting-visualizations-and-programmatic-control-office-scripts"></a>Mise en forme, visualisations et contrôle programmatique : Office scripts

Nous vous recommandons Office scripts lorsque vos besoins vont au-delà de l’importation et de la transformation des données.

Presque tout ce que vous pouvez faire manuellement via l Excel’interface utilisateur peut être fait à l’Office scripts. Ils sont parfaits pour appliquer une mise en forme cohérente aux workbooks. Les scripts créent des graphiques, des tableaux croisés dynamiques, des formes, des images et d’autres visualisations de feuille de calcul. Les scripts vous donnent également un contrôle précis sur les positions, tailles, couleurs et autres attributs de ces visualisations.

L’inclusion de code TypeScript vous donne un haut degré de personnalisation. Une logique de contrôle programmatique telle `if...else` que des instructions rend votre script robuste. Cela vous permet de faire des opérations telles que lire des données de manière conditionnable sans vous appuyer sur des formules Excel complexes, ou d’analyser le workbook pour y voir des modifications inattendues avant de modifier le manuel.

La mise en forme peut être appliquée avec Power Query via Excel [modèles.](https://templates.office.com/power-query-tutorial-tm11414620) Toutefois, les modèles sont mis à jour au niveau individuel ou de l’organisation, tandis que Office Scripts offrent un contrôle d’accès plus granulaire.

## <a name="power-automate-integrations"></a>Power Automate’intégrations

Office scripts offrent davantage d’options pour Power Automate’intégration. Les scripts sont adaptés à vos solutions. Vous définissez [l’entrée et la sortie](../develop/power-automate-integration.md#data-transfer-in-flows-for-scripts)du script, afin qu’il fonctionne avec tout autre connecteur ou toute autre donnée dans le flux. La capture d’écran suivante montre un exemple Power Automate flux de données qui transmet les données d’une carte adaptative Teams à un script Office.

:::image type="content" source="../images/scenario-task-reminders-last-flow-step.png" alt-text="Capture d’écran shows the Excel Online (Business) connector in the flow designer. Le connecteur utilise l’action de script Exécuter pour prendre une entrée à partir d Teams carte adaptative et la fournir à un script.":::

Power Query est utilisé dans le [connecteur SQL Server](https://powerquery.microsoft.com/flow/) Power Automate’alimentation. Les [données transform à l’aide de l’action](/connectors/sql/#transform-data-using-power-query) Power Query vous permet de créer une requête dans Power Automate. Bien qu’il s’agit d’un outil puissant à utiliser avec SQL Server, il limite Power Query à cette source d’entrée, comme illustré dans la capture d’écran de flux suivante.

:::image type="content" source="../images/power-query-flow-option.png" alt-text="Capture d’écran shows the SQL Server connector in the flow designer. Le connecteur utilise les données transform à l’aide de l’action Power Query.":::

## <a name="platform-dependencies"></a>Dépendances de plateforme

Office Scripts est actuellement disponible uniquement pour les Excel sur le Web. Power Query n’est actuellement disponible que pour Excel sur ordinateur de bureau. Les deux peuvent être utilisés par Power Automate, ce qui permet au flux de fonctionner avec Excel de travail stockés dans OneDrive.

## <a name="see-also"></a>Voir aussi

- [Portail Power Query](https://powerquery.microsoft.com/)
- [Power Query avec Excel](https://powerquery.microsoft.com/excel/)
- [Exécuter Office scripts avec Power Automate](../develop/power-automate-integration.md)
