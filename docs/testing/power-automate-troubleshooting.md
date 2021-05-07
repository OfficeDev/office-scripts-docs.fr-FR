---
title: Informations de dépannage pour les Power Automate avec Office scripts
description: Astuces, les informations de plateforme et les problèmes connus avec l’intégration entre Office scripts et Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: bcfedb8db88d74f16e46c604121bceff3c7c7382
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232647"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Informations de dépannage pour les Power Automate avec Office scripts

Power Automate vous permet d’Office l’automatisation des scripts au niveau suivant. Toutefois, comme Power Automate exécute des scripts en votre nom dans des sessions Excel indépendantes, il existe quelques points importants à noter.

> [!TIP]
> Si vous débutez avec Office Scripts avec Power Automate, commencez par Exécuter [des scripts Office](../develop/power-automate-integration.md) avec Power Automate pour en savoir plus sur les plateformes.

## <a name="avoid-using-relative-references"></a>Éviter d’utiliser des références relatives

Power Automate exécute votre script dans le Excel de travail choisi en votre nom. Le workbook peut être fermé lorsque cela se produit. Toute API qui repose sur l’état actuel de l’utilisateur, par exemple, peut se comporter différemment dans `Workbook.getActiveWorksheet` Power Automate. Cela est dû au fait que les API sont basées sur une position relative de l’affichage ou du curseur de l’utilisateur et que cette référence n’existe pas dans Power Automate flux.

Certaines API de référence relative envoient des erreurs Power Automate. D’autres ont un comportement par défaut qui implique l’état d’un utilisateur. Lors de la conception de vos scripts, n’oubliez pas d’utiliser des références absolues pour les feuilles de calcul et les plages. Cela permet à votre Power Automate un flux cohérent, même si les feuilles de calcul sont réorganiser.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Méthodes de script qui échouent lors de l’Power Automate flux

Les méthodes suivantes lancent une erreur et échouent lorsqu’elles sont appelées à partir d’un script dans Power Automate flux.

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

Les méthodes suivantes utilisent un comportement par défaut, à la place de l’état actuel de n’importe quel utilisateur.

| Classe | Méthode | Power Automate comportement |
|--|--|--|
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Renvoie la première feuille de calcul du manuel ou la feuille de calcul actuellement activée par la `Worksheet.activate` méthode. |
| [Feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marque la feuille de calcul comme feuille de calcul active à des fins de `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Sélectionner des classes avec le contrôle de navigateur de fichiers

Lors de la création **de l’étape** d’Power Automate script d’un flux d’Power Automate, vous devez sélectionner le workbook qui fait partie du flux. Utilisez le navigateur de fichiers pour sélectionner votre classer, au lieu de taper manuellement le nom du classer.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="L’Power Automate exécuter l’action de script affichant l’option Afficher le navigateur de fichier du s picker":::

Pour plus de contexte sur la limitation Power Automate et une discussion sur les solutions de contournement potentielles pour la sélection dynamique de workbooks, voir ce thread dans le microsoft [Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Différences de fuseau horaire

Excel fichiers n’ont pas d’emplacement ou de fuseau horaire inhérents. Chaque fois qu’un utilisateur ouvre le manuel, sa session utilise le fuseau horaire local de cet utilisateur pour les calculs de date. Power Automate utilise toujours l’UTC.

Si votre script utilise des dates ou des heures, il peut y avoir des différences de comportement lorsque le script est testé localement par rapport au moment où il est exécuté Power Automate. Power Automate vous permet de convertir, de mettre en forme et d’ajuster les temps. Voir [Utilisation](https://flow.microsoft.com/blog/working-with-dates-and-times/) de dates et d’heures dans vos flux pour obtenir des instructions sur l’utilisation de ces fonctions dans Power Automate and [ `main` Parameters: Passing data to a script](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) to learn how to provide that time information for the script.

## <a name="see-also"></a>Voir aussi

- [Dépannage de Office Scripts](troubleshooting.md)
- [Exécuter Office scripts avec Power Automate](../develop/power-automate-integration.md)
- [Excel Documentation de référence sur le connecteur en ligne (Entreprise)](/connectors/excelonlinebusiness/)
