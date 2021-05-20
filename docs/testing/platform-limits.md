---
title: Limites et exigences de la plate-forme avec Office scripts
description: Limites de ressources et prise en charge du navigateur pour les scripts Office lorsqu’ils sont utilisés avec Excel sur le Web
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7e81aaf2f96faeb67c815814fe3b7f1795651318
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545580"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites et exigences de la plate-forme avec Office scripts

Il existe certaines limitations de plate-forme dont vous devez être conscient lors du développement Office scripts. Cet article détaille le support du navigateur et les limites de données Office scripts pour Excel sur le Web.

## <a name="browser-support"></a>Prise en charge du navigateur

Office Scripts fonctionnent dans n’importe quel navigateur [qui prend en charge Office sur le Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Toutefois, certaines fonctionnalités JavaScript ne sont pas prises en charge dans Internet Explorer 11 (IE 11). Toutes les fonctionnalités [introduites dans ES6 ou](https://www.w3schools.com/Js/js_es6.asp) plus tard ne fonctionnera pas avec IE 11. Si les personnes de votre organisation utilisent toujours ce navigateur, assurez-vous de tester vos scripts dans cet environnement lorsque vous les partagez.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies tiers

Votre navigateur a besoin de cookies tiers activés pour afficher **l’onglet Automate** dans Excel sur le Web. Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché. Si vous utilisez une session de navigateur privée, vous devrez peut-être ré-activer ce paramètre à chaque fois.

> [!NOTE]
> Certains navigateurs se réfèrent à ce paramètre comme « tous les cookies », au lieu de « cookies tiers ».

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instructions pour ajuster les paramètres des cookies dans les navigateurs populaires

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites des données

Il y a des limites à la quantité Excel données peuvent être transférées à la fois et au nombre de transactions Power Automate peuvent être effectuées.

### <a name="excel"></a>Excel

Excel sur le Web les limites suivantes lors des appels au cahier de travail par le biais d’un script :

- Les demandes et les réponses sont limitées **à 5 Mo**.
- Une portée est limitée à **cinq millions de cellules.**

Si vous rencontrez des erreurs lorsque vous traitez de grands ensembles de données, essayez d’utiliser plusieurs plages plus petites au lieu de plus grandes plages. Par exemple, consultez l’exemple [Écrire un grand ensemble de données.](../resources/samples/write-large-dataset.md) Vous pouvez également utiliser des API comme [Range.getSpecialCells pour](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) cibler des cellules spécifiques au lieu de grandes plages.

### <a name="power-automate"></a>Power Automate

Lors de l Office scripts avec Power Automate, chaque utilisateur est limité **à 400 appels à l’action Run Script par jour**. Cette limite se réinitialise à 00h00 UTC.

La Power Automate plate-forme a également des limitations d’utilisation, qui peuvent être trouvées dans les articles suivants:

- [Limites et configuration dans Power Automate](/power-automate/limits-and-config)
- [Problèmes et limitations connus pour le connecteur Excel en ligne (Business)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a>Voir aussi

- [Scripts de Office dépannage](troubleshooting.md)
- [Annuler les effets des scripts Office texte](undo.md)
- [Améliorez les performances de vos scripts Office’argent](../develop/web-client-performance.md)
- [Scripting Fundamentals for Office Scripts in Excel sur le Web](../develop/scripting-fundamentals.md)
