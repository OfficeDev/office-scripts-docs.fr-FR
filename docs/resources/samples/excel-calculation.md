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
# <a name="manage-calculation-mode-in-excel"></a>Gérer le mode de calcul dans Excel

Cet exemple montre comment utiliser le [mode de calcul](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) et calculer des méthodes dans Excel sur le Web à l’aide Office scripts. Vous pouvez essayer le script sur n’importe Excel fichier.

## <a name="scenario"></a>Scénario

Dans Excel sur le Web, le mode de calcul d’un fichier peut être contrôlé par programme à l’aide d’API. Les actions suivantes sont possibles à l’aide Office scripts.

1. Obtenir le mode de calcul.
1. Définissez le mode de calcul.
1. Calculez Excel formules pour les fichiers qui sont définies en mode manuel (également appelé recalcul).

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

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/iw6O8QH01CI).
