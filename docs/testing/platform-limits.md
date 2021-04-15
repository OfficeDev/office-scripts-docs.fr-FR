---
title: Limites et exigences de plateforme avec les scripts Office
description: Limites de ressources et prise en charge des navigateurs pour les scripts Office lorsqu'ils sont utilisés avec Excel sur le web
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: ef733562fb3caa8261fbbd8382923927a46cb7d4
ms.sourcegitcommit: 5ca286615a11d282e3f80023d22d36a039800eed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2021
ms.locfileid: "51689765"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="ddf27-103">Limites et exigences de plateforme avec les scripts Office</span><span class="sxs-lookup"><span data-stu-id="ddf27-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="ddf27-104">Certaines limitations de plateforme sont à prendre en compte lors du développement de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="ddf27-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="ddf27-105">Cet article décrit en détail la prise en charge du navigateur et les limites de données pour les scripts Office pour Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="ddf27-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="ddf27-106">Prise en charge du navigateur</span><span class="sxs-lookup"><span data-stu-id="ddf27-106">Browser support</span></span>

<span data-ttu-id="ddf27-107">Les scripts Office fonctionnent dans n'importe quel navigateur [qui prend en charge Office pour le web.](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452)</span><span class="sxs-lookup"><span data-stu-id="ddf27-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="ddf27-108">Toutefois, certaines fonctionnalités JavaScript ne sont pas pris en charge dans Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="ddf27-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="ddf27-109">Toutes les fonctionnalités introduites dans [ES6](https://www.w3schools.com/Js/js_es6.asp) ou une ultérieure ne fonctionnent pas avec IE 11.</span><span class="sxs-lookup"><span data-stu-id="ddf27-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="ddf27-110">Si les membres de votre organisation utilisent toujours ce navigateur, n'oubliez pas de tester vos scripts dans cet environnement lors de leur partage.</span><span class="sxs-lookup"><span data-stu-id="ddf27-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="ddf27-111">Cookies tiers</span><span class="sxs-lookup"><span data-stu-id="ddf27-111">Third-party cookies</span></span>

<span data-ttu-id="ddf27-112">Votre navigateur a besoin de cookies tiers activés pour afficher **l'onglet Automatiser** dans Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="ddf27-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="ddf27-113">Vérifiez les paramètres de votre navigateur si l'onglet n'est pas affiché.</span><span class="sxs-lookup"><span data-stu-id="ddf27-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="ddf27-114">Si vous utilisez une session de navigateur privé, vous devrez peut-être activer ce paramètre à chaque fois.</span><span class="sxs-lookup"><span data-stu-id="ddf27-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="ddf27-115">Certains navigateurs font référence à ce paramètre en tant que « tous les cookies », au lieu de « cookies tiers ».</span><span class="sxs-lookup"><span data-stu-id="ddf27-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="ddf27-116">Instructions d'ajustement des paramètres de cookie dans les navigateurs populaires</span><span class="sxs-lookup"><span data-stu-id="ddf27-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="ddf27-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="ddf27-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="ddf27-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="ddf27-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="ddf27-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="ddf27-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="ddf27-120">Safari</span><span class="sxs-lookup"><span data-stu-id="ddf27-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="ddf27-121">Limites des données</span><span class="sxs-lookup"><span data-stu-id="ddf27-121">Data limits</span></span>

<span data-ttu-id="ddf27-122">Il existe des limites sur la quantité de données Excel qui peuvent être transférées en même temps et sur le nombre de transactions Power Automate individuelles qui peuvent être effectuées.</span><span class="sxs-lookup"><span data-stu-id="ddf27-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="ddf27-123">Excel</span><span class="sxs-lookup"><span data-stu-id="ddf27-123">Excel</span></span>

<span data-ttu-id="ddf27-124">Excel pour le web présente les limitations suivantes lors de l'appel au workbook par le biais d'un script :</span><span class="sxs-lookup"><span data-stu-id="ddf27-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="ddf27-125">Les demandes et réponses sont limitées à **5 Mo.**</span><span class="sxs-lookup"><span data-stu-id="ddf27-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="ddf27-126">Une plage est limitée à **cinq millions de cellules.**</span><span class="sxs-lookup"><span data-stu-id="ddf27-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="ddf27-127">Si vous rencontrez des erreurs lorsque vous traitez des jeux de données volumineux, essayez d'utiliser plusieurs plages plus petites plutôt que des plages plus grandes.</span><span class="sxs-lookup"><span data-stu-id="ddf27-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="ddf27-128">Vous pouvez également utiliser des API telles [que Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) pour cibler des cellules spécifiques au lieu de grandes plages.</span><span class="sxs-lookup"><span data-stu-id="ddf27-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="ddf27-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="ddf27-129">Power Automate</span></span>

<span data-ttu-id="ddf27-130">Lorsque vous utilisez Office Scripts avec Power Automate, chaque utilisateur est limité à **400** appels à l'action Exécuter le script par jour.</span><span class="sxs-lookup"><span data-stu-id="ddf27-130">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="ddf27-131">Cette limite est réinitialisée à 00h00 UTC.</span><span class="sxs-lookup"><span data-stu-id="ddf27-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="ddf27-132">La plateforme Power Automate présente également des limitations d'utilisation, qui sont présentes dans les articles suivants :</span><span class="sxs-lookup"><span data-stu-id="ddf27-132">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="ddf27-133">Limites et configuration dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="ddf27-133">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="ddf27-134">Problèmes connus et limitations pour le connecteur Excel Online (entreprise)</span><span class="sxs-lookup"><span data-stu-id="ddf27-134">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="ddf27-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ddf27-135">See also</span></span>

- [<span data-ttu-id="ddf27-136">Dépannage de Office Scripts</span><span class="sxs-lookup"><span data-stu-id="ddf27-136">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="ddf27-137">Annuler les effets d’un script Office</span><span class="sxs-lookup"><span data-stu-id="ddf27-137">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="ddf27-138">Améliorer les performances de vos scripts Office</span><span class="sxs-lookup"><span data-stu-id="ddf27-138">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="ddf27-139">Principes de base des scripts pour les scripts Office dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="ddf27-139">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
