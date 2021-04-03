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
# <a name="manage-calculation-mode-in-excel"></a>Gérer le mode de calcul dans Excel

Cet exemple montre comment utiliser le [mode de calcul](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) et calculer des méthodes dans Excel sur le web à l’aide de Scripts Office. Vous pouvez essayer le script sur n’importe quel fichier Excel.

## <a name="scenario"></a>Scénario

Dans Excel sur le web, le mode de calcul d’un fichier peut être contrôlé par programme à l’aide d’API. Les actions suivantes sont possibles à l’aide des scripts Office.

1. Obtenir le mode de calcul.
1. Définissez le mode de calcul.
1. Calculer des formules Excel pour les fichiers qui sont définies en mode manuel (également appelé recalcul).

## <a name="sample-code-control-calculation-mode"></a>Exemple de code : mode de calcul de contrôle

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

## <a name="training-video-manage-calculation-mode"></a>Vidéo de formation : gérer le mode de calcul

[![Regardez une vidéo pas à pas sur la gestion du mode de calcul dans Excel sur le web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Vidéo pas à pas sur la gestion du mode de calcul dans Excel sur le web")
