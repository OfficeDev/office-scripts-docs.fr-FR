---
title: Gérer le mode de calcul dans Excel
description: Découvrez comment utiliser Office Scripts pour gérer le mode de calcul dans Excel sur le web.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 0239437c7b52dca1fd8d1a4fc66bab7965cbd91a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571281"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="fd16d-103">Gérer le mode de calcul dans Excel</span><span class="sxs-lookup"><span data-stu-id="fd16d-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="fd16d-104">Cet exemple montre comment utiliser le [mode de calcul](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) et calculer des méthodes dans Excel sur le web à l’aide de Scripts Office.</span><span class="sxs-lookup"><span data-stu-id="fd16d-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="fd16d-105">Vous pouvez essayer le script sur n’importe quel fichier Excel.</span><span class="sxs-lookup"><span data-stu-id="fd16d-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="fd16d-106">Scénario</span><span class="sxs-lookup"><span data-stu-id="fd16d-106">Scenario</span></span>

<span data-ttu-id="fd16d-107">Dans Excel sur le web, le mode de calcul d’un fichier peut être contrôlé par programme à l’aide d’API.</span><span class="sxs-lookup"><span data-stu-id="fd16d-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="fd16d-108">Les actions suivantes sont possibles à l’aide des scripts Office.</span><span class="sxs-lookup"><span data-stu-id="fd16d-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="fd16d-109">Obtenir le mode de calcul.</span><span class="sxs-lookup"><span data-stu-id="fd16d-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="fd16d-110">Définissez le mode de calcul.</span><span class="sxs-lookup"><span data-stu-id="fd16d-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="fd16d-111">Calculer des formules Excel pour les fichiers qui sont définies en mode manuel (également appelé recalcul).</span><span class="sxs-lookup"><span data-stu-id="fd16d-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="fd16d-112">Exemple de code : mode de calcul de contrôle</span><span class="sxs-lookup"><span data-stu-id="fd16d-112">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="fd16d-113">Vidéo de formation : gérer le mode de calcul</span><span class="sxs-lookup"><span data-stu-id="fd16d-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="fd16d-114">[![Regardez une vidéo pas à pas sur la gestion du mode de calcul dans Excel sur le web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Vidéo pas à pas sur la gestion du mode de calcul dans Excel sur le web")</span><span class="sxs-lookup"><span data-stu-id="fd16d-114">[![Watch step-by-step video on how to manage calculation mode in Excel on the web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Step-by-step video on how to manage calculation mode in Excel on the web")</span></span>
