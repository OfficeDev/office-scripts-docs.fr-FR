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
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="4f85a-103">Limites et exigences de la plateforme avec les scripts Office</span><span class="sxs-lookup"><span data-stu-id="4f85a-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="4f85a-104">Il existe certaines limitations de plateforme dont vous devez être conscient lors du développement de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="4f85a-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="4f85a-105">Cet article décrit la prise en charge du navigateur et les limites de données pour les scripts Office pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="4f85a-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="4f85a-106">Prise en charge du navigateur</span><span class="sxs-lookup"><span data-stu-id="4f85a-106">Browser support</span></span>

<span data-ttu-id="4f85a-107">Les scripts Office fonctionnent dans n’importe quel navigateur qui [prend en charge Office pour le Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="4f85a-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="4f85a-108">Toutefois, certaines fonctionnalités JavaScript ne sont pas prises en charge dans Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="4f85a-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="4f85a-109">Toutes les fonctionnalités introduites dans [ES6 ou version ultérieure](https://www.w3schools.com/Js/js_es6.asp) ne fonctionneront pas avec Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="4f85a-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="4f85a-110">Si les membres de votre organisation continuent d’utiliser ce navigateur, veillez à tester vos scripts dans cet environnement lors de leur partage.</span><span class="sxs-lookup"><span data-stu-id="4f85a-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="4f85a-111">Cookies tiers</span><span class="sxs-lookup"><span data-stu-id="4f85a-111">Third-party cookies</span></span>

<span data-ttu-id="4f85a-112">Votre navigateur a besoin de cookies tiers activés pour afficher l’onglet **automatiser** dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="4f85a-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="4f85a-113">Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché.</span><span class="sxs-lookup"><span data-stu-id="4f85a-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="4f85a-114">Si vous utilisez une session de navigateur privé, vous devrez peut-être réactiver ce paramètre à chaque fois.</span><span class="sxs-lookup"><span data-stu-id="4f85a-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="4f85a-115">Certains navigateurs se réfèrent à ce paramètre comme « tous les cookies », au lieu de « cookies tiers ».</span><span class="sxs-lookup"><span data-stu-id="4f85a-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="4f85a-116">Instructions pour la modification des paramètres de cookies dans les navigateurs populaires</span><span class="sxs-lookup"><span data-stu-id="4f85a-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="4f85a-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="4f85a-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="4f85a-118">Edge</span><span class="sxs-lookup"><span data-stu-id="4f85a-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="4f85a-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="4f85a-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="4f85a-120">Safari</span><span class="sxs-lookup"><span data-stu-id="4f85a-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="4f85a-121">Limites des données</span><span class="sxs-lookup"><span data-stu-id="4f85a-121">Data limits</span></span>

<span data-ttu-id="4f85a-122">Il existe des limites quant à la quantité de données Excel pouvant être transférées en une seule fois, ainsi que le nombre de transactions d’automate de puissance individuelles pouvant être effectuées.</span><span class="sxs-lookup"><span data-stu-id="4f85a-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="4f85a-123">Excel</span><span class="sxs-lookup"><span data-stu-id="4f85a-123">Excel</span></span>

<span data-ttu-id="4f85a-124">Excel pour le Web présente les limitations suivantes lors de l’appel du classeur via un script :</span><span class="sxs-lookup"><span data-stu-id="4f85a-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="4f85a-125">Les demandes et les réponses sont limitées à **5 Mo**.</span><span class="sxs-lookup"><span data-stu-id="4f85a-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="4f85a-126">Une plage est limitée à **5 millions cellules**.</span><span class="sxs-lookup"><span data-stu-id="4f85a-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="4f85a-127">Si vous rencontrez des erreurs lorsque vous traitez des jeux de données volumineux, essayez d’utiliser plusieurs plages plus petites au lieu de plages plus grandes.</span><span class="sxs-lookup"><span data-stu-id="4f85a-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="4f85a-128">Vous pouvez également des API comme [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) pour cibler des cellules spécifiques au lieu de grandes plages.</span><span class="sxs-lookup"><span data-stu-id="4f85a-128">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="4f85a-129">Power Automate</span><span class="sxs-lookup"><span data-stu-id="4f85a-129">Power Automate</span></span>

<span data-ttu-id="4f85a-130">Lorsque vous utilisez des scripts Office avec automate Power, vous êtes limité à **200 appels par jour**.</span><span class="sxs-lookup"><span data-stu-id="4f85a-130">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="4f85a-131">Cette limite est rétablie à 12:00 AM UTC.</span><span class="sxs-lookup"><span data-stu-id="4f85a-131">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="4f85a-132">La plateforme de gestion de l’alimentation automatique présente également des limitations d’utilisation, qui se trouvent dans les articles [Limits and configuration in Power Automated](/power-automate/limits-and-config).</span><span class="sxs-lookup"><span data-stu-id="4f85a-132">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="4f85a-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4f85a-133">See also</span></span>

- [<span data-ttu-id="4f85a-134">Dépannage de Office Scripts</span><span class="sxs-lookup"><span data-stu-id="4f85a-134">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="4f85a-135">Annuler les effets d’un script Office</span><span class="sxs-lookup"><span data-stu-id="4f85a-135">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="4f85a-136">Améliorer les performances de vos scripts Office</span><span class="sxs-lookup"><span data-stu-id="4f85a-136">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="4f85a-137">Scripts de base pour les scripts Office dans Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="4f85a-137">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
