---
title: Limites et exigences de plateforme avec les scripts Office
description: Limites de ressources et prise en charge des navigateurs pour les scripts Office lorsqu’ils sont utilisés avec Excel sur le web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 93307b6204f409f26c77b5ead33188205d5c4b4d
ms.sourcegitcommit: 5bde455b06ee2ed007f3e462d8ad485b257774ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50837264"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites et exigences de plateforme avec les scripts Office

Certaines limitations de plateforme sont à prendre en compte lors du développement de scripts Office. Cet article décrit en détail la prise en charge du navigateur et les limites de données pour les scripts Office pour Excel sur le web.

## <a name="browser-support"></a>Prise en charge du navigateur

Les scripts Office fonctionnent dans n’importe quel navigateur [qui prend en charge Office pour le web.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452) Toutefois, certaines fonctionnalités JavaScript ne sont pas pris en charge dans Internet Explorer 11 (IE 11). Toutes les fonctionnalités introduites dans [ES6](https://www.w3schools.com/Js/js_es6.asp) ou une ultérieure ne fonctionnent pas avec IE 11. Si les membres de votre organisation utilisent toujours ce navigateur, n’oubliez pas de tester vos scripts dans cet environnement lors de leur partage.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies tiers

Votre navigateur a besoin de cookies tiers activés pour afficher **l’onglet Automatiser** dans Excel sur le web. Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché. Si vous utilisez une session de navigateur privé, vous devrez peut-être activer ce paramètre à chaque fois.

> [!NOTE]
> Certains navigateurs font référence à ce paramètre en tant que « tous les cookies », au lieu de « cookies tiers ».

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instructions d’ajustement des paramètres de cookie dans les navigateurs populaires

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites des données

Il existe des limites sur la quantité de données Excel qui peuvent être transférées en même temps et sur le nombre de transactions Power Automate individuelles qui peuvent être effectuées.

### <a name="excel"></a>Excel

Excel pour le web présente les limitations suivantes lors de l’appel au workbook par le biais d’un script :

- Les demandes et réponses sont limitées à **5 Mo.**
- Une plage est limitée à **cinq millions de cellules.**

Si vous rencontrez des erreurs lorsque vous traitez des jeux de données volumineux, essayez d’utiliser plusieurs plages plus petites plutôt que des plages plus grandes. Vous pouvez également utiliser des API [telles que Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) pour cibler des cellules spécifiques au lieu de grandes plages.

### <a name="power-automate"></a>Power Automate

Lorsque vous utilisez Office Scripts avec Power Automate, chaque utilisateur est limité à **200 appels par jour.** Cette limite est réinitialisée à 00h00 UTC.

La plateforme Power Automate présente également des limitations d’utilisation, qui sont présentes dans les articles suivants :

- [Limites et configuration dans Power Automate](/power-automate/limits-and-config)
- [Problèmes connus et limitations pour le connecteur Excel Online (entreprise)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Voir aussi

- [Dépannage de Office Scripts](troubleshooting.md)
- [Annuler les effets d’un script Office](undo.md)
- [Améliorer les performances de vos scripts Office](../develop/web-client-performance.md)
- [Principes de base des scripts pour les scripts Office dans Excel sur le web](../develop/scripting-fundamentals.md)
