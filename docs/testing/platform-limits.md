---
title: Limites et exigences de plateforme avec Office scripts
description: Limites de ressources et prise en charge du navigateur pour Office scripts lorsqu’ils sont utilisés avec Excel sur le Web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 8b7afa02f73476e6e98f231a7a7162ad87607b37
ms.sourcegitcommit: 9d00ee1c11cdf897410e5232692ee985f01ee098
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53772357"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites et exigences de plateforme avec Office scripts

Certaines limitations de plateforme sont à prendre en compte lors du développement de Office scripts. Cet article décrit en détail la prise en charge du navigateur et les limites de données pour Office scripts pour Excel sur le Web.

## <a name="browser-support"></a>Prise en charge du navigateur

Office Les scripts fonctionnent dans n’importe quel navigateur qui [prend en charge Office sur le Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Toutefois, certaines fonctionnalités JavaScript ne sont pas pris en charge dans Internet Explorer 11 (IE 11). Toutes les fonctionnalités introduites dans [ES6](https://www.w3schools.com/Js/js_es6.asp) ou une ultérieure ne fonctionnent pas avec IE 11. Si les membres de votre organisation utilisent toujours ce navigateur, n’oubliez pas de tester vos scripts dans cet environnement lors de leur partage.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies tiers

Votre navigateur a besoin de cookies tiers activés pour afficher l’onglet **Automatiser** dans Excel sur le Web. Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché. Si vous utilisez une session de navigateur privé, vous devrez peut-être activer ce paramètre à chaque fois.

> [!NOTE]
> Certains navigateurs font référence à ce paramètre en tant que « tous les cookies », au lieu de « cookies tiers ».

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instructions d’ajustement des paramètres de cookie dans les navigateurs populaires

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites des données

Il existe des limites sur le nombre Excel données peuvent être transférées en même temps et sur le nombre Power Automate transactions peuvent être effectuées.

### <a name="excel"></a>Excel

Excel sur le Web présente les limitations suivantes lors de l’appel au workbook via un script :

- Les demandes et réponses sont limitées à **5 Mo.**
- Une plage est limitée à **cinq millions de cellules.**

Si vous rencontrez des erreurs lorsque vous traitez des jeux de données volumineux, essayez d’utiliser plusieurs plages plus petites plutôt que des plages plus grandes. Pour obtenir un exemple, [consultez l’exemple Écrire un jeu de données](../resources/samples/write-large-dataset.md) de grande taille. Vous pouvez également utiliser des API telles [que Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getSpecialCells_cellType__cellValueType_) pour cibler des cellules spécifiques au lieu de grandes plages.

### <a name="power-automate"></a>Power Automate

Lorsque vous utilisez Office scripts avec Power Automate, chaque utilisateur est limité à **400 appels** à l’action Exécuter le script par jour. Cette limite est réinitialisée à 00h00 UTC.

La plateforme Power Automate a également des limitations d’utilisation, qui sont présentes dans les articles suivants :

- [Limites et configuration dans Power Automate](/power-automate/limits-and-config)
- [Problèmes connus et limitations pour le connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Voir aussi

- [Résoudre les problèmes Office scripts](troubleshooting.md)
- [Annuler les effets des scripts Office](undo.md)
- [Améliorer les performances de vos scripts Office de gestion](../develop/web-client-performance.md)
- [Principes de base des scripts Office scripts dans Excel sur le Web](../develop/scripting-fundamentals.md)
