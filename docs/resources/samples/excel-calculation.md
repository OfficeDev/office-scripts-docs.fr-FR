---
title: Gérer le mode de calcul dans Excel
description: Découvrez comment utiliser Office Scripts pour gérer le mode de calcul dans Excel sur le Web.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: a60fddc91b3a8f124a44722d0d75e6e9f239351d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285912"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="e6728-103">Gérer le mode de calcul dans Excel</span><span class="sxs-lookup"><span data-stu-id="e6728-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="e6728-104">Cet exemple montre comment utiliser le [mode de calcul](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) et calculer des méthodes dans Excel sur le Web à l’aide Office scripts.</span><span class="sxs-lookup"><span data-stu-id="e6728-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="e6728-105">Vous pouvez essayer le script sur n’importe Excel fichier.</span><span class="sxs-lookup"><span data-stu-id="e6728-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="e6728-106">Scénario</span><span class="sxs-lookup"><span data-stu-id="e6728-106">Scenario</span></span>

<span data-ttu-id="e6728-107">Le recalcul des workbooks avec un grand nombre de formules peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="e6728-107">Workbooks with large numbers of formulas can take a while to recalculate.</span></span> <span data-ttu-id="e6728-108">Au lieu de laisser Excel contrôler le moment où les calculs ont lieu, vous pouvez les gérer dans le cadre de votre script.</span><span class="sxs-lookup"><span data-stu-id="e6728-108">Rather than letting Excel control when calculations happen, you can manage them as part of your script.</span></span> <span data-ttu-id="e6728-109">Cela permet d’améliorer les performances dans certains scénarios.</span><span class="sxs-lookup"><span data-stu-id="e6728-109">This will help with performance in certain scenarios.</span></span>

<span data-ttu-id="e6728-110">L’exemple de script définit le mode de calcul sur manuel.</span><span class="sxs-lookup"><span data-stu-id="e6728-110">The sample script sets the calculation mode to manual.</span></span> <span data-ttu-id="e6728-111">Cela signifie que le workbook ne recalcule les formules que lorsque le script l’indique (ou que vous calculez manuellement via [l’interface utilisateur).](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)</span><span class="sxs-lookup"><span data-stu-id="e6728-111">This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span></span> <span data-ttu-id="e6728-112">Le script affiche ensuite le mode de calcul actuel et recalcule entièrement le workbook entier.</span><span class="sxs-lookup"><span data-stu-id="e6728-112">The script then displays the current calculation mode and fully recalculates the entire workbook.</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="e6728-113">Exemple de code : mode de calcul de contrôle</span><span class="sxs-lookup"><span data-stu-id="e6728-113">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="e6728-114">Vidéo de formation : gérer le mode de calcul</span><span class="sxs-lookup"><span data-stu-id="e6728-114">Training video: Manage calculation mode</span></span>

<span data-ttu-id="e6728-115">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/iw6O8QH01CI).</span><span class="sxs-lookup"><span data-stu-id="e6728-115">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
