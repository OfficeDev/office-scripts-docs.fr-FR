---
title: Limites et exigences de plateforme avec Office scripts
description: Limites de ressources et prise en charge du navigateur pour Office scripts lorsqu’ils sont utilisés avec Excel sur le Web.
ms.date: 01/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 385248e5c62ed3dbf2827105b3097ef27e5187a7
ms.sourcegitcommit: b84d4c8dd31335e4e39b0da6ad25fd528cb9d8f3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62462501"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites et exigences de plateforme avec Office scripts

Certaines limitations de plateforme sont à prendre en compte lors du développement de Office scripts. Cet article décrit en détail la prise en charge du navigateur et les limites de données pour Office scripts pour Excel sur le Web.

## <a name="browser-support"></a>Prise en charge du navigateur

Office scripts fonctionnent dans n’importe quel navigateur qui [prend en charge Office sur le Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Toutefois, certaines fonctionnalités JavaScript ne sont pas pris en charge dans Internet Explorer 11 (IE 11). Toutes les fonctionnalités introduites dans [ES6](https://www.w3schools.com/Js/js_es6.asp) ou une ultérieure ne fonctionnent pas avec IE 11. Si les membres de votre organisation utilisent toujours ce navigateur, n’oubliez pas de tester vos scripts dans cet environnement lors de leur partage.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies tiers

Votre navigateur a besoin de cookies tiers activés pour afficher **l’onglet Automatiser** dans Excel sur le Web. Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché. Si vous utilisez une session de navigateur privé, vous devrez peut-être activer ce paramètre à chaque fois.

> [!NOTE]
> Certains navigateurs font référence à ce paramètre en tant que « tous les cookies », au lieu de « cookies tiers ».

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instructions d’ajustement des paramètres de cookie dans les navigateurs populaires

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites des données

Il existe des limites sur le nombre Excel données peuvent être transférées en même temps et sur le nombre Power Automate transactions peuvent être effectuées.

### <a name="excel"></a>Excel

Excel sur le Web présente les limitations suivantes lors de l’appel au workbook par le biais d’un script :

- Les demandes et les réponses sont limitées **à 5 Mo**.
- Une plage est limitée à **cinq millions de cellules**.

Si vous rencontrez des erreurs lorsque vous traitez des jeux de données volumineux, essayez d’utiliser plusieurs plages plus petites plutôt que des plages plus grandes. Pour obtenir un exemple, [consultez l’exemple Écrire un jeu de données de grande](../resources/samples/write-large-dataset.md) taille. Vous pouvez également utiliser des API telles [que Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) pour cibler des cellules spécifiques au lieu de grandes plages.

### <a name="power-automate"></a>Power Automate

Lorsque vous utilisez Office scripts avec Power Automate, chaque utilisateur est limité à **1 600 appels par jour à l’action Exécuter le script**. Cette limite est réinitialisée à 00h00 UTC.

La plateforme Power Automate a également des limitations d’utilisation, qui sont présentes dans les articles suivants.

- [Limites et configuration dans Power Automate](/power-automate/limits-and-config)
- [Problèmes connus et limitations pour le connecteur Excel Online (Entreprise)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Si vous avez un script de longue durée, n’ignorez pas le délai [d’Power Automate 120 secondes](/power-automate/limits-and-config#timeout). Vous devez optimiser votre [script](../develop/web-client-performance.md) ou fractionner votre automatisation Excel en plusieurs scripts.

## <a name="see-also"></a>Voir aussi

- [Résoudre les problèmes Office scripts](troubleshooting.md)
- [Annuler les effets des scripts Office](undo.md)
- [Améliorer les performances de vos scripts Office de gestion](../develop/web-client-performance.md)
- [Principes de base des scripts Office scripts dans Excel sur le Web](../develop/scripting-fundamentals.md)
