---
title: Gérer le mode de calcul dans Excel
description: Découvrez comment utiliser Office Scripts pour gérer le mode de calcul dans Excel sur le Web.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 34a14874197ffda8487df5e450e3dcab980f7ed5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232451"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="31419-103">Gérer le mode de calcul dans Excel</span><span class="sxs-lookup"><span data-stu-id="31419-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="31419-104">Cet exemple montre comment utiliser le [mode de calcul](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) et calculer des méthodes dans Excel sur le Web à l’aide Office scripts.</span><span class="sxs-lookup"><span data-stu-id="31419-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="31419-105">Vous pouvez essayer le script sur n’importe Excel fichier.</span><span class="sxs-lookup"><span data-stu-id="31419-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="31419-106">Scénario</span><span class="sxs-lookup"><span data-stu-id="31419-106">Scenario</span></span>

<span data-ttu-id="31419-107">Dans Excel sur le Web, le mode de calcul d’un fichier peut être contrôlé par programme à l’aide d’API.</span><span class="sxs-lookup"><span data-stu-id="31419-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="31419-108">Les actions suivantes sont possibles à l’aide Office scripts.</span><span class="sxs-lookup"><span data-stu-id="31419-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="31419-109">Obtenir le mode de calcul.</span><span class="sxs-lookup"><span data-stu-id="31419-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="31419-110">Définissez le mode de calcul.</span><span class="sxs-lookup"><span data-stu-id="31419-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="31419-111">Calculez Excel formules pour les fichiers qui sont définies en mode manuel (également appelé recalcul).</span><span class="sxs-lookup"><span data-stu-id="31419-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="31419-112">Exemple de code : mode de calcul de contrôle</span><span class="sxs-lookup"><span data-stu-id="31419-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="31419-113">Vidéo de formation : gérer le mode de calcul</span><span class="sxs-lookup"><span data-stu-id="31419-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="31419-114">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="31419-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
