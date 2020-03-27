---
title: Différences entre les scripts Office et les compléments Office
description: Les différences de comportement et d’API entre les scripts Office et les compléments Office.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978714"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="72c1a-103">Différences entre les scripts Office et les compléments Office</span><span class="sxs-lookup"><span data-stu-id="72c1a-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="72c1a-104">Les compléments Office et les scripts Office ont beaucoup de choses en commun.</span><span class="sxs-lookup"><span data-stu-id="72c1a-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="72c1a-105">Ils proposent tous les deux un contrôle automatique d’un classeur Excel via l' `Excel` espace de noms de l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="72c1a-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="72c1a-106">Toutefois, les scripts Office sont plus limités dans leur étendue.</span><span class="sxs-lookup"><span data-stu-id="72c1a-106">However, Office Scripts are more limited in their scope.</span></span>

![Diagramme à quatre quadrants montrant les zones ciblées pour différentes solutions d’extensibilité Office.](../images/office-programmability-diagram.png)

<span data-ttu-id="72c1a-109">Les scripts Office sont exécutés jusqu’à la fin avec une pression manuelle ou une étape de l' [automate d’alimentation](https://flow.microsoft.com/), tandis que les compléments Office persistent lorsque leurs volets Office sont ouverts.</span><span class="sxs-lookup"><span data-stu-id="72c1a-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="72c1a-110">Cela signifie que les compléments peuvent conserver l’État pendant une session, tandis que les scripts Office ne gèrent pas un état interne entre les exécutions.</span><span class="sxs-lookup"><span data-stu-id="72c1a-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="72c1a-111">Si vous constatez que votre extension Excel doit dépasser les fonctionnalités de la plateforme de script, consultez la [documentation relative aux compléments Office](/office/dev/add-ins) pour en savoir plus sur les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="72c1a-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="72c1a-112">Le reste de cet article décrit les principales différences entre les compléments Office et les scripts Office.</span><span class="sxs-lookup"><span data-stu-id="72c1a-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="72c1a-113">Prise en charge de la plateforme</span><span class="sxs-lookup"><span data-stu-id="72c1a-113">Platform Support</span></span>

<span data-ttu-id="72c1a-114">Les compléments Office sont multiplateformes.</span><span class="sxs-lookup"><span data-stu-id="72c1a-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="72c1a-115">Elles fonctionnent sur des plateformes de bureau Windows, Mac, iOS et Web et fournissent la même expérience sur chacun d’eux.</span><span class="sxs-lookup"><span data-stu-id="72c1a-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="72c1a-116">Toutes les exceptions sont indiquées dans la documentation de l’API individuelle.</span><span class="sxs-lookup"><span data-stu-id="72c1a-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="72c1a-117">Les scripts Office sont actuellement uniquement pris en charge par Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="72c1a-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="72c1a-118">Toutes les opérations d’enregistrement, de modification et d’exécution sont réalisées sur la plateforme Web.</span><span class="sxs-lookup"><span data-stu-id="72c1a-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="72c1a-119">API</span><span class="sxs-lookup"><span data-stu-id="72c1a-119">APIs</span></span>

<span data-ttu-id="72c1a-120">Les scripts Office prennent en charge la plupart des API JavaScript pour Excel, ce qui signifie qu’il existe un grand nombre de fonctionnalités qui se chevauchent entre les deux plateformes.</span><span class="sxs-lookup"><span data-stu-id="72c1a-120">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="72c1a-121">Il existe deux exceptions : les événements et les API communes.</span><span class="sxs-lookup"><span data-stu-id="72c1a-121">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="72c1a-122">Événements</span><span class="sxs-lookup"><span data-stu-id="72c1a-122">Events</span></span>

<span data-ttu-id="72c1a-123">Les scripts Office ne prennent pas en charge les [événements](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="72c1a-123">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="72c1a-124">Chaque script exécute le code dans une méthode `main` unique, puis se termine.</span><span class="sxs-lookup"><span data-stu-id="72c1a-124">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="72c1a-125">Il ne réactive pas lorsque des événements sont déclenchés et, par conséquent, ne peut pas enregistrer d’événements.</span><span class="sxs-lookup"><span data-stu-id="72c1a-125">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="72c1a-126">API communes</span><span class="sxs-lookup"><span data-stu-id="72c1a-126">Common APIs</span></span>

<span data-ttu-id="72c1a-127">Les scripts Office ne peuvent pas utiliser des [API communes](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="72c1a-127">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="72c1a-128">Si vous avez besoin d’une authentification, de fenêtres de boîtes de dialogue ou d’autres fonctionnalités qui sont uniquement prises en charge par des API communes, vous aurez probablement besoin de créer un complément Office au lieu d’un script Office.</span><span class="sxs-lookup"><span data-stu-id="72c1a-128">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="72c1a-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="72c1a-129">See also</span></span>

- [<span data-ttu-id="72c1a-130">Office Scripts dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="72c1a-130">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="72c1a-131">Différences entre les scripts Office et les macros VBA</span><span class="sxs-lookup"><span data-stu-id="72c1a-131">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="72c1a-132">Dépannage de Office Scripts</span><span class="sxs-lookup"><span data-stu-id="72c1a-132">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="72c1a-133">Créer un complément de volet de tâches Excel</span><span class="sxs-lookup"><span data-stu-id="72c1a-133">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
