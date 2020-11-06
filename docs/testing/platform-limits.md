---
title: Limites et exigences de la plateforme avec les scripts Office
description: Limites de ressources et prise en charge de navigateur pour les scripts Office lorsqu’ils sont utilisés avec Excel sur le Web
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930077"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites et exigences de la plateforme avec les scripts Office

Il existe certaines limitations de plateforme dont vous devez être conscient lors du développement de scripts Office. Cet article décrit la prise en charge du navigateur et les limites de données pour les scripts Office pour Excel sur le Web.

## <a name="browser-support"></a>Prise en charge du navigateur

Les scripts Office fonctionnent dans n’importe quel navigateur qui [prend en charge Office pour le Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Toutefois, certaines fonctionnalités JavaScript ne sont pas prises en charge dans Internet Explorer 11 (IE 11). Toutes les fonctionnalités introduites dans [ES6 ou version ultérieure](https://www.w3schools.com/Js/js_es6.asp) ne fonctionneront pas avec Internet Explorer 11. Si les membres de votre organisation continuent d’utiliser ce navigateur, veillez à tester vos scripts dans cet environnement lors de leur partage.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies tiers

Votre navigateur a besoin de cookies tiers activés pour afficher l’onglet **automatiser** dans Excel sur le Web. Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché. Si vous utilisez une session de navigateur privé, vous devrez peut-être réactiver ce paramètre à chaque fois.

> [!NOTE]
> Certains navigateurs se réfèrent à ce paramètre comme « tous les cookies », au lieu de « cookies tiers ».

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instructions pour la modification des paramètres de cookies dans les navigateurs populaires

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites des données

Il existe des limites quant à la quantité de données Excel pouvant être transférées en une seule fois, ainsi que le nombre de transactions d’automate de puissance individuelles pouvant être effectuées.

### <a name="excel"></a>Excel

Excel pour le Web présente les limitations suivantes lors de l’appel du classeur via un script :

- Les demandes et les réponses sont limitées à **5 Mo**.
- Une plage est limitée à **5 millions cellules**.

Si vous rencontrez des erreurs lorsque vous traitez des jeux de données volumineux, essayez d’utiliser plusieurs plages plus petites au lieu de plages plus grandes. Vous pouvez également des API comme [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) pour cibler des cellules spécifiques au lieu de grandes plages.

### <a name="power-automate"></a>Power Automate

Lorsque vous utilisez des scripts Office avec automate Power, vous êtes limité à **200 appels par jour**. Cette limite est rétablie à 12:00 AM UTC.

La plateforme de gestion de l’alimentation automatique présente également des limitations d’utilisation, qui se trouvent dans les articles [Limits and configuration in Power Automated](/power-automate/limits-and-config).

## <a name="see-also"></a>Voir aussi

- [Dépannage de Office Scripts](troubleshooting.md)
- [Annuler les effets d’un script Office](undo.md)
- [Améliorer les performances de vos scripts Office](../develop/web-client-performance.md)
- [Scripts de base pour les scripts Office dans Excel sur le Web](../develop/scripting-fundamentals.md)
