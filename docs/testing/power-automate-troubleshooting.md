---
title: Dépannage Office scripts en cours d’exécution Power Automate
description: Astuces, informations de plate-forme, et les problèmes connus avec l’intégration entre Office scripts et Power Automate.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e26378051c764d97b4e8d748abc85fbe095c7b03
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545568"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Dépannage Office scripts en cours d’exécution Power Automate

Power Automate vous permet de passer Office l’automatisation du script au niveau suivant. Toutefois, comme Power Automate exécute des scripts en votre nom dans des sessions Excel, il y a quelques choses importantes à noter.

> [!TIP]
> Si vous commencez tout juste à utiliser Office scripts avec Power Automate, veuillez commencer par exécuter des scripts Office avec des Power Automate pour en apprendre davantage sur les [plateformes.](../develop/power-automate-integration.md)

## <a name="avoid-relative-references"></a>Éviter les références relatives

Power Automate exécute votre script dans le cahier de travail Excel choisi en votre nom. Le cahier de travail peut être fermé lorsque cela se produit. Toute API qui s’appuie sur l’état actuel de l’utilisateur, `Workbook.getActiveWorksheet` par exemple, peut se comporter différemment dans Power Automate. C’est parce que les API sont basées sur une position relative de la vue ou du curseur de l’utilisateur et que cette référence n’existe pas dans un flux Power Automate utilisateur.

Certaines API de référence relative jettent des erreurs dans Power Automate. D’autres ont un comportement par défaut qui implique l’état d’un utilisateur. Lors de la conception de vos scripts, assurez-vous d’utiliser des références absolues pour les feuilles de travail et les plages. Cela rend votre Power Automate de débit constant, même si les feuilles de travail sont réarrangées.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Méthodes de script qui échouent lorsqu’elles sont Power Automate flux

Les méthodes suivantes lanceront une erreur et échoueront lorsqu’elles sont appelées à partir d’un script dans Power Automate flux.

| Classe | Méthode |
|--|--|
| [Graphique](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Méthodes de script avec un comportement par défaut dans Power Automate flux

Les méthodes suivantes utilisent un comportement par défaut, au lieu de l’état actuel de tout utilisateur.

| Classe | Méthode | Power Automate comportement |
|--|--|--|
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Renvoie soit la première feuille de travail dans le cahier de travail, soit la feuille de travail actuellement activée par la `Worksheet.activate` méthode. |
| [Feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marque la feuille de travail comme la feuille de travail active aux fins de `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Sélectionnez des cahiers de travail avec le contrôle du navigateur de fichiers

Lors de la **création de l’étape** de script Exécuter Power Automate flux de flux, vous devez sélectionner quel manuel fait partie du flux. Utilisez le navigateur de fichiers pour sélectionner votre cahier de travail, au lieu de taper manuellement le nom du livre de travail.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="L’action Power Automate exécuter le script montrant l’option de navigateur de fichiers Show Picker":::

Pour plus de contexte sur la limitation Power Automate et une discussion des solutions de contournement potentielles pour la sélection dynamique des cahiers de travail, voir [ce fil dans le microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Différences de fuseau horaire

Excel fichiers n’ont pas d’emplacement ou de fuseau horaire inhérent. Chaque fois qu’un utilisateur ouvre le cahier de travail, sa session utilise le fuseau horaire local de cet utilisateur pour les calculs de date. Power Automate utilise toujours UTC.

Si votre script utilise des dates ou des heures, il peut y avoir des différences comportementales lorsque le script est testé localement par rapport à quand il est exécuté à travers Power Automate. Power Automate vous permet de convertir, formater et ajuster les temps. Voir [Travailler avec les dates et les heures à l’intérieur de vos](https://flow.microsoft.com/blog/working-with-dates-and-times/) flux pour obtenir des instructions sur la façon d’utiliser ces fonctions dans Power Automate et paramètres : [ `main` transmettez des](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) données à un script pour apprendre à fournir ces informations de temps pour le script.

## <a name="see-also"></a>Voir aussi

- [Scripts de Office dépannage](troubleshooting.md)
- [Exécutez Office scripts avec Power Automate](../develop/power-automate-integration.md)
- [Excel Documentation de référence connecteur en ligne (Business)](/connectors/excelonlinebusiness/)
