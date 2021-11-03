---
title: Résoudre les problèmes Office scripts en cours d’exécution dans Power Automate
description: Astuces, les informations de plateforme et les problèmes connus avec l’intégration entre Office scripts et Power Automate.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 028c34003a6f6b00c9afc67450b249b938d445fb
ms.sourcegitcommit: 634ad2061e683ae1032c1e0b55b00ac577adc34f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/03/2021
ms.locfileid: "60725628"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>Résoudre les problèmes Office scripts en cours d’exécution dans Power Automate

Power Automate vous permet d’Office l’automatisation des scripts au niveau suivant. Toutefois, comme Power Automate exécute des scripts en votre nom dans des sessions Excel indépendantes, il existe quelques points importants à noter.

> [!TIP]
> Si vous débutez avec Office Scripts avec Power Automate, commencez par Exécuter [des scripts Office](../develop/power-automate-integration.md) avec Power Automate pour en savoir plus sur les plateformes.

## <a name="avoid-relative-references"></a>Éviter les références relatives

Power Automate exécute votre script dans le Excel de travail choisi en votre nom. Le workbook peut être fermé lorsque cela se produit. Toute API qui repose sur l’état actuel de l’utilisateur, par exemple, peut se comporter différemment dans `Workbook.getActiveWorksheet` Power Automate. Cela est dû au fait que les API sont basées sur une position relative de l’affichage ou du curseur de l’utilisateur et que cette référence n’existe pas dans Power Automate flux.

Certaines API de référence relative envoient des erreurs Power Automate. D’autres ont un comportement par défaut qui implique l’état d’un utilisateur. Lors de la conception de vos scripts, n’oubliez pas d’utiliser des références absolues pour les feuilles de calcul et les plages. Cela permet à votre Power Automate un flux cohérent, même si les feuilles de calcul sont réorganiser.

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>Méthodes de script qui échouent lors de l’Power Automate flux

Les méthodes suivantes envoient une erreur et échouent lorsqu’elles sont appelées à partir d’un script dans Power Automate flux.

| Classe | Méthode |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Méthodes de script avec un comportement par défaut dans Power Automate flux

Les méthodes suivantes utilisent un comportement par défaut, à la place de l’état actuel de n’importe quel utilisateur.

| Classe | Méthode | Power Automate comportement |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Renvoie la première feuille de calcul du manuel ou la feuille de calcul actuellement activée par la `Worksheet.activate` méthode. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marque la feuille de calcul en tant que feuille de calcul active à des fins de `Workbook.getActiveWorksheet` . |

## <a name="data-refresh-not-supported-in-power-automate"></a>L’actualisation des données n’est pas prise en charge dans Power Automate

Office Les scripts ne peuvent pas actualiser les données lorsqu’ils sont exécutés Power Automate. Méthodes telles que `PivotTable.refresh` ne rien faire lorsqu’elles sont appelées dans un flux. En outre, Power Automate ne déclenche pas d’actualisation des données pour les formules qui utilisent des liens de workbook.

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>Méthodes de script qui ne font rien lorsqu’elles sont Power Automate flux

Les méthodes suivantes ne font rien dans un script lorsqu’elles sont appelées Power Automate. Elles sont toujours correctement renvoy es et ne lancent pas d’erreurs.

| Classe | Méthode |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Classeur](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Feuille de calcul](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>Sélectionner des classes avec le contrôle de navigateur de fichiers

Lors de la création **de l’étape** d’Power Automate script d’un flux d’Power Automate, vous devez sélectionner le workbook qui fait partie du flux. Utilisez le navigateur de fichiers pour sélectionner votre classez, au lieu de taper manuellement le nom du classer.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="L’Power Automate exécuter une action de script montrant l’option Afficher le navigateur de fichier du s picker.":::

Pour plus de contexte sur la limitation Power Automate et une discussion sur les solutions de contournement potentielles pour la sélection dynamique de workbooks, voir ce thread dans le microsoft [Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="pass-entire-arrays-as-script-parameters"></a>Passer des tableaux entiers en tant que paramètres de script

Power Automate permet aux utilisateurs de transmettre des tableaux à des connecteurs en tant que variable ou en tant qu’éléments simples dans le tableau. La valeur par défaut consiste à transmettre des éléments simples, ce qui crée le tableau dans le flux. Pour les scripts ou autres connecteurs qui prennent des  tableaux entiers en tant qu’arguments, vous devez sélectionner le bouton Basculer pour entrer l’intégralité du tableau pour passer le tableau en tant qu’objet complet. Ce bouton se trouve dans le coin supérieur droit de chaque champ d’entrée de paramètre de tableau.

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="Bouton à basculer pour entrer un tableau entier dans une zone d’entrée de champ de contrôle.":::

## <a name="time-zone-differences"></a>Différences de fuseau horaire

Excel fichiers n’ont pas d’emplacement ou de fuseau horaire inhérents. Chaque fois qu’un utilisateur ouvre le manuel, sa session utilise le fuseau horaire local de cet utilisateur pour les calculs de date. Power Automate utilise toujours l’UTC.

Si votre script utilise des dates ou des heures, il peut y avoir des différences de comportement lorsque le script est testé localement par rapport au moment où il est exécuté Power Automate. Power Automate vous permet de convertir, de mettre en forme et d’ajuster les temps. Voir [Utilisation](https://flow.microsoft.com/blog/working-with-dates-and-times/) des dates et heures à l’intérieur de vos flux pour obtenir des instructions sur l’utilisation de ces fonctions dans Power Automate and [ `main` Parameters: Pass data to a script to](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) learn how to provide that time information for the script.

## <a name="script-parameter-fields-or-returned-output-not-appearing-in-power-automate"></a>Les champs de paramètre de script ou la sortie renvoyée n’apparaissent pas dans Power Automate

Il existe deux raisons pour lesquelles les paramètres ou les données renvoyées d’un script ne sont pas reflétés avec précision dans le Power Automate de flux.

- La signature du script (paramètres ou valeur de retour) a changé depuis l’ajout **du connecteur Excel Business (Online).**
- La signature de script utilise des types non pris en place. Vérifiez vos types par rapport aux listes sous les [paramètres](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) et retournez les [sections](../develop/power-automate-integration.md#return-data-from-a-script) de Run [Office Scripts avec Power Automate](../develop/power-automate-integration.md) article.

La signature d’un script est stockée avec le **connecteur Excel Business (Online)** lors de sa création. Supprimez l’ancien connecteur et créez-en un pour obtenir les derniers paramètres et valeurs de retour pour l’action **de script Exécuter.**

## <a name="see-also"></a>Voir aussi

- [Résoudre les problèmes Office scripts](troubleshooting.md)
- [Exécuter Office scripts avec Power Automate](../develop/power-automate-integration.md)
- [Excel Documentation de référence sur le connecteur en ligne (Entreprise)](/connectors/excelonlinebusiness/)
