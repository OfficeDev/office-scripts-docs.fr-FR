---
title: Informations de dépannage pour Power Automate avec les scripts Office
description: Conseils, informations sur la plateforme et problèmes connus avec l'intégration entre Office Scripts et Power Automate.
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: 59f4cd8b3476c2ee2a1a862f136173a543ba8a15
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755006"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Informations de dépannage pour Power Automate avec les scripts Office

Power Automate vous permet d'aller au niveau suivant de votre automatisation de script Office. Toutefois, étant donné que Power Automate exécute des scripts en votre nom dans des sessions Excel indépendantes, il existe quelques points importants à noter.

> [!TIP]
> Si vous commencez à utiliser les scripts Office avec Power Automate, commencez par exécuter [des scripts Office](../develop/power-automate-integration.md) avec Power Automate pour en savoir plus sur les plateformes.

## <a name="avoid-using-relative-references"></a>Éviter d'utiliser des références relatives

Power Automate exécute votre script dans le workbook Excel choisi en votre nom. Le workbook peut être fermé lorsque cela se produit. Toute API qui repose sur l'état actuel de l'utilisateur, par exemple, peut se comporter `Workbook.getActiveWorksheet` différemment dans Power Automate. Cela est dû au fait que les API sont basées sur une position relative de l'affichage ou du curseur de l'utilisateur et que cette référence n'existe pas dans un flux Power Automate.

Certaines API de référence relative permettent de créer des erreurs dans Power Automate. D'autres ont un comportement par défaut qui implique l'état d'un utilisateur. Lors de la conception de vos scripts, n'oubliez pas d'utiliser des références absolues pour les feuilles de calcul et les plages. Votre flux Power Automate est ainsi cohérent, même si les feuilles de calcul sont réorganiser.

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>Méthodes de script qui échouent lors de l'exécuter des flux Power Automate

Les méthodes suivantes lancent une erreur et échouent lorsqu'elles sont appelées à partir d'un script dans un flux Power Automate.

| Classe | Méthode |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Méthodes de script avec un comportement par défaut dans les flux Power Automate

Les méthodes suivantes utilisent un comportement par défaut, à la place de l'état actuel de n'importe quel utilisateur.

| Classe | Méthode | Comportement de Power Automate |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | Renvoie la première feuille de calcul du manuel ou la feuille de calcul actuellement activée par la `Worksheet.activate` méthode. |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | Marque la feuille de calcul comme feuille de calcul active à des fins de `Workbook.getActiveWorksheet` . |

## <a name="select-workbooks-with-the-file-browser-control"></a>Sélectionner des classes avec le contrôle de navigateur de fichiers

Lorsque vous **construisez l'étape du script** Exécuter d'un flux Power Automate, vous devez sélectionner le workbook qui fait partie du flux. Utilisez le navigateur de fichiers pour sélectionner votre classez, au lieu de taper manuellement le nom du classer.

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="Action de script Exécuter Power Automate montrant l'option Afficher le navigateur du fichier du s picker.":::

Pour plus de contexte sur la limitation De Power Automate et une discussion sur les solutions de contournement potentielles pour la sélection dynamique de workbooks, voir ce thread dans la [communauté Microsoft Power Automate](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).

## <a name="time-zone-differences"></a>Différences de fuseau horaire

Les fichiers Excel n'ont pas d'emplacement ou de fuseau horaire inhérents. Chaque fois qu'un utilisateur ouvre le manuel, sa session utilise le fuseau horaire local de cet utilisateur pour les calculs de date. Power Automate utilise toujours l'UTC.

Si votre script utilise des dates ou des heures, il peut y avoir des différences de comportement lorsque le script est testé localement par rapport au moment où il est exécuté via Power Automate. Power Automate vous permet de convertir, de mettre en forme et d'ajuster les temps. Consultez l'utilisation de [dates](https://flow.microsoft.com/blog/working-with-dates-and-times/) et d'heures dans vos flux pour obtenir des instructions sur l'utilisation de ces fonctions dans Power Automate et [ `main` paramètres](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) : transmettre des données à un script pour découvrir comment fournir ces informations d'heure pour le script.

## <a name="see-also"></a>Voir aussi

- [Dépannage de Office Scripts](troubleshooting.md)
- [Exécuter des scripts Office avec Power Automate](../develop/power-automate-integration.md)
- [Documentation de référence sur le connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness/)
