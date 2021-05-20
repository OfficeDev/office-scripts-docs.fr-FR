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
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="d9fec-103">Limites et exigences de la plate-forme avec Office scripts</span><span class="sxs-lookup"><span data-stu-id="d9fec-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="d9fec-104">Il existe certaines limitations de plate-forme dont vous devez être conscient lors du développement Office scripts.</span><span class="sxs-lookup"><span data-stu-id="d9fec-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="d9fec-105">Cet article détaille le support du navigateur et les limites de données Office scripts pour Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="d9fec-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="d9fec-106">Prise en charge du navigateur</span><span class="sxs-lookup"><span data-stu-id="d9fec-106">Browser support</span></span>

<span data-ttu-id="d9fec-107">Office Scripts fonctionnent dans n’importe quel navigateur [qui prend en charge Office sur le Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="d9fec-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="d9fec-108">Toutefois, certaines fonctionnalités JavaScript ne sont pas prises en charge dans Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="d9fec-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="d9fec-109">Toutes les fonctionnalités [introduites dans ES6 ou](https://www.w3schools.com/Js/js_es6.asp) plus tard ne fonctionnera pas avec IE 11.</span><span class="sxs-lookup"><span data-stu-id="d9fec-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="d9fec-110">Si les personnes de votre organisation utilisent toujours ce navigateur, assurez-vous de tester vos scripts dans cet environnement lorsque vous les partagez.</span><span class="sxs-lookup"><span data-stu-id="d9fec-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a><span data-ttu-id="d9fec-111">Cookies tiers</span><span class="sxs-lookup"><span data-stu-id="d9fec-111">Third-party cookies</span></span>

<span data-ttu-id="d9fec-112">Votre navigateur a besoin de cookies tiers activés pour afficher **l’onglet Automate** dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="d9fec-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="d9fec-113">Vérifiez les paramètres de votre navigateur si l’onglet n’est pas affiché.</span><span class="sxs-lookup"><span data-stu-id="d9fec-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="d9fec-114">Si vous utilisez une session de navigateur privée, vous devrez peut-être ré-activer ce paramètre à chaque fois.</span><span class="sxs-lookup"><span data-stu-id="d9fec-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="d9fec-115">Certains navigateurs se réfèrent à ce paramètre comme « tous les cookies », au lieu de « cookies tiers ».</span><span class="sxs-lookup"><span data-stu-id="d9fec-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a><span data-ttu-id="d9fec-116">Instructions pour ajuster les paramètres des cookies dans les navigateurs populaires</span><span class="sxs-lookup"><span data-stu-id="d9fec-116">Instructions for adjusting cookie settings in popular browsers</span></span>

- [<span data-ttu-id="d9fec-117">Chrome</span><span class="sxs-lookup"><span data-stu-id="d9fec-117">Chrome</span></span>](https://support.google.com/chrome/answer/95647)
- [<span data-ttu-id="d9fec-118">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="d9fec-118">Edge</span></span>](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [<span data-ttu-id="d9fec-119">Firefox</span><span class="sxs-lookup"><span data-stu-id="d9fec-119">Firefox</span></span>](https://support.mozilla.org/kb/disable-third-party-cookies)
- [<span data-ttu-id="d9fec-120">Safari</span><span class="sxs-lookup"><span data-stu-id="d9fec-120">Safari</span></span>](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a><span data-ttu-id="d9fec-121">Limites des données</span><span class="sxs-lookup"><span data-stu-id="d9fec-121">Data limits</span></span>

<span data-ttu-id="d9fec-122">Il y a des limites à la quantité Excel données peuvent être transférées à la fois et au nombre de transactions Power Automate peuvent être effectuées.</span><span class="sxs-lookup"><span data-stu-id="d9fec-122">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="d9fec-123">Excel</span><span class="sxs-lookup"><span data-stu-id="d9fec-123">Excel</span></span>

<span data-ttu-id="d9fec-124">Excel sur le Web les limites suivantes lors des appels au cahier de travail par le biais d’un script :</span><span class="sxs-lookup"><span data-stu-id="d9fec-124">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="d9fec-125">Les demandes et les réponses sont limitées **à 5 Mo**.</span><span class="sxs-lookup"><span data-stu-id="d9fec-125">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="d9fec-126">Une portée est limitée à **cinq millions de cellules.**</span><span class="sxs-lookup"><span data-stu-id="d9fec-126">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="d9fec-127">Si vous rencontrez des erreurs lorsque vous traitez de grands ensembles de données, essayez d’utiliser plusieurs plages plus petites au lieu de plus grandes plages.</span><span class="sxs-lookup"><span data-stu-id="d9fec-127">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="d9fec-128">Par exemple, consultez l’exemple [Écrire un grand ensemble de données.](../resources/samples/write-large-dataset.md)</span><span class="sxs-lookup"><span data-stu-id="d9fec-128">For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample.</span></span> <span data-ttu-id="d9fec-129">Vous pouvez également utiliser des API comme [Range.getSpecialCells pour](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) cibler des cellules spécifiques au lieu de grandes plages.</span><span class="sxs-lookup"><span data-stu-id="d9fec-129">You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="d9fec-130">Power Automate</span><span class="sxs-lookup"><span data-stu-id="d9fec-130">Power Automate</span></span>

<span data-ttu-id="d9fec-131">Lors de l Office scripts avec Power Automate, chaque utilisateur est limité **à 400 appels à l’action Run Script par jour**.</span><span class="sxs-lookup"><span data-stu-id="d9fec-131">When using Office Scripts with Power Automate, each user is limited to **400 calls to the Run Script action per day**.</span></span> <span data-ttu-id="d9fec-132">Cette limite se réinitialise à 00h00 UTC.</span><span class="sxs-lookup"><span data-stu-id="d9fec-132">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="d9fec-133">La Power Automate plate-forme a également des limitations d’utilisation, qui peuvent être trouvées dans les articles suivants:</span><span class="sxs-lookup"><span data-stu-id="d9fec-133">The Power Automate platform also has usage limitations, which can be found in the following articles:</span></span>

- [<span data-ttu-id="d9fec-134">Limites et configuration dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="d9fec-134">Limits and configuration in Power Automate</span></span>](/power-automate/limits-and-config)
- [<span data-ttu-id="d9fec-135">Problèmes et limitations connus pour le connecteur Excel en ligne (Business)</span><span class="sxs-lookup"><span data-stu-id="d9fec-135">Known issues and limitations for the Excel Online (Business) connector</span></span>](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## <a name="see-also"></a><span data-ttu-id="d9fec-136">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d9fec-136">See also</span></span>

- [<span data-ttu-id="d9fec-137">Scripts de Office dépannage</span><span class="sxs-lookup"><span data-stu-id="d9fec-137">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="d9fec-138">Annuler les effets des scripts Office texte</span><span class="sxs-lookup"><span data-stu-id="d9fec-138">Undo the effects of Office Scripts</span></span>](undo.md)
- [<span data-ttu-id="d9fec-139">Améliorez les performances de vos scripts Office’argent</span><span class="sxs-lookup"><span data-stu-id="d9fec-139">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="d9fec-140">Scripting Fundamentals for Office Scripts in Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="d9fec-140">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
