---
title: Différences entre les scripts Office et les compléments Office
description: Les différences de comportement et d’API entre les scripts Office et les compléments Office.
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700229"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="6c0ac-103">Différences entre les scripts Office et les compléments Office</span><span class="sxs-lookup"><span data-stu-id="6c0ac-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="6c0ac-104">Les compléments Office et les scripts Office ont beaucoup de choses en commun.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="6c0ac-105">Ils proposent tous les deux un contrôle automatique d’un classeur Excel via l' `Excel` espace de noms de l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="6c0ac-106">Toutefois, les scripts Office sont plus limités dans leur étendue.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-106">However, Office Scripts are more limited in their scope.</span></span>

<span data-ttu-id="6c0ac-107">Exécution des scripts Office avec une pression manuelle, tandis que les compléments Office s’appuient sur l’interaction de l’utilisateur et sont persistants pendant l’utilisation du classeur.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-107">Office Scripts run to completion with a manual button press, whereas Office Add-ins rely on user interaction and persist while the workbook is in use.</span></span> <span data-ttu-id="6c0ac-108">Si vous constatez que votre extension Excel doit dépasser les fonctionnalités de la plateforme de script, consultez la [documentation relative aux compléments Office](/office/dev/add-ins) pour en savoir plus sur les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-108">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="6c0ac-109">Le reste de cet article décrit les principales différences entre les compléments Office et les scripts Office.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-109">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="6c0ac-110">Prise en charge de la plateforme</span><span class="sxs-lookup"><span data-stu-id="6c0ac-110">Platform Support</span></span>

<span data-ttu-id="6c0ac-111">Les compléments Office sont multiplateformes.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-111">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="6c0ac-112">Elles fonctionnent sur des plateformes de bureau Windows, Mac, iOS et Web et fournissent la même expérience sur chacun d’eux.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-112">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="6c0ac-113">Toutes les exceptions sont indiquées dans la documentation de l’API individuelle.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-113">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="6c0ac-114">Les scripts Office sont actuellement uniquement pris en charge par Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-114">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="6c0ac-115">Toutes les opérations d’enregistrement, de modification et d’exécution sont réalisées sur la plateforme Web.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-115">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="6c0ac-116">API</span><span class="sxs-lookup"><span data-stu-id="6c0ac-116">APIs</span></span>

<span data-ttu-id="6c0ac-117">Les scripts Office prennent en charge la plupart des API JavaScript pour Excel, ce qui signifie qu’il existe un grand nombre de fonctionnalités qui se chevauchent entre les deux plateformes.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-117">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="6c0ac-118">Il existe deux exceptions : les événements et les API communes.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-118">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="6c0ac-119">Événements</span><span class="sxs-lookup"><span data-stu-id="6c0ac-119">Events</span></span>

<span data-ttu-id="6c0ac-120">Les scripts Office ne prennent pas en charge les [événements](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="6c0ac-120">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="6c0ac-121">Chaque script exécute le code dans une méthode `main` unique, puis se termine.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-121">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="6c0ac-122">Il ne réactive pas lorsque des événements sont déclenchés et, par conséquent, ne peut pas enregistrer d’événements.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-122">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="6c0ac-123">API communes</span><span class="sxs-lookup"><span data-stu-id="6c0ac-123">Common APIs</span></span>

<span data-ttu-id="6c0ac-124">Les scripts Office ne peuvent pas utiliser des [API communes](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="6c0ac-124">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="6c0ac-125">Si vous avez besoin d’une authentification, de fenêtres de boîtes de dialogue ou d’autres fonctionnalités qui sont uniquement prises en charge par des API communes, vous aurez probablement besoin de créer un complément Office au lieu d’un script Office.</span><span class="sxs-lookup"><span data-stu-id="6c0ac-125">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="6c0ac-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6c0ac-126">See also</span></span>

- [<span data-ttu-id="6c0ac-127">Scripts Office dans Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="6c0ac-127">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="6c0ac-128">Résolution des problèmes liés aux scripts Office</span><span class="sxs-lookup"><span data-stu-id="6c0ac-128">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="6c0ac-129">Créer un complément de volet de tâches Excel</span><span class="sxs-lookup"><span data-stu-id="6c0ac-129">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)