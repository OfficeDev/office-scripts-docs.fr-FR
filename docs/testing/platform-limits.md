---
title: Limites et exigences de la plateforme avec les scripts Office
description: Limites de ressources et prise en charge des navigateurs pour les scripts Office lorsqu’ils sont utilisés avec Excel sur le Web.
ms.date: 11/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 764d1eddaf303a941a098ec1d3f3056d63e8693f
ms.sourcegitcommit: 7cadf2b637bf62874e43b6e595286101816662aa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/09/2022
ms.locfileid: "68891245"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites et exigences de la plateforme avec les scripts Office

Il existe certaines limitations de plateforme que vous devez connaître lors du développement de scripts Office. Cet article détaille la prise en charge des navigateurs et les limites de données pour les scripts Office pour Excel sur le Web.

## <a name="browser-support"></a>Prise en charge du navigateur

Les scripts Office fonctionnent dans n’importe quel navigateur qui [prend en charge Office sur le Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). Toutefois, certaines fonctionnalités JavaScript ne sont pas prises en charge dans Internet Explorer 11 (IE 11). Les fonctionnalités introduites dans [ES6 ou version ultérieure](https://www.w3schools.com/Js/js_es6.asp) ne fonctionnent pas avec Internet Explorer 11. Si les membres de votre organisation utilisent toujours ce navigateur, veillez à tester vos scripts dans cet environnement lorsque vous les partagez.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies tiers

Votre navigateur doit activer les cookies tiers pour afficher l’onglet **Automatiser** dans Excel sur le Web. Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché. Si vous utilisez une session de navigateur privé, vous devrez peut-être réactiver ce paramètre à chaque fois.

> [!NOTE]
> Certains navigateurs font référence à ce paramètre en tant que « tous les cookies », au lieu de « cookies tiers ».

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instructions pour ajuster les paramètres des cookies dans les navigateurs populaires

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Microsoft Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites des données

Il existe des limites quant à la quantité de données Excel pouvant être transférées simultanément et au nombre de transactions Power Automate individuelles pouvant être effectuées.

### <a name="excel"></a>Excel

Excel sur le Web présente les limitations suivantes lors des appels au classeur via un script :

- Les demandes et les réponses sont limitées à **5 Mo**.
- Une plage est limitée à **cinq millions de cellules**.

Si vous rencontrez des erreurs lors du traitement de jeux de données volumineux, essayez d’utiliser plusieurs plages plus petites au lieu de plages plus grandes. Pour obtenir un exemple, consultez l’exemple [Écrire un jeu de données volumineux](../resources/samples/write-large-dataset.md) . Vous pouvez également utiliser des API telles que [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) pour cibler des cellules spécifiques au lieu de grandes plages.

Vous trouverez des limites Excel qui ne sont pas spécifiques aux scripts Office dans l’article [Spécifications et limites d’Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).

### <a name="power-automate"></a>Power Automate

Lorsque vous utilisez des scripts Office avec Power Automate, chaque utilisateur est limité à **1 600 appels à l’action Exécuter le script par jour**. Cette limite est réinitialisée à 00:00 UTC.

La plateforme Power Automate présente également des limitations d’utilisation, que vous trouverez dans les articles suivants.

- [Limites et configuration dans Power Automate](/power-automate/limits-and-config)
- [Problèmes connus et limitations du connecteur Excel Online (Business)](/connectors/excelonlinebusiness/#known-issues-and-limitations)

> [!NOTE]
> Si vous disposez d’un script de longue durée, tenez compte du [délai d’expiration de 120 secondes pour les opérations Power Automate synchrones](/power-automate/limits-and-config#timeout). Vous devez [optimiser votre script](../develop/web-client-performance.md) ou fractionner votre automatisation Excel en plusieurs scripts.

## <a name="see-also"></a>Voir aussi

- [Spécifications et limites relatives à Excel](https://support.microsoft.com/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
- [Résoudre les problèmes liés aux scripts Office](troubleshooting.md)
- [Annuler les effets des scripts Office](undo.md)
- [Améliorer les performances de vos scripts Office](../develop/web-client-performance.md)
