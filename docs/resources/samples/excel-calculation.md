---
title: Gérer le mode de calcul dans Excel
description: Découvrez comment utiliser Office Scripts pour gérer le mode de calcul dans Excel sur le Web.
ms.date: 05/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: 32ed55f47106c7ff2dadb21aca7fce71ff7d2b3d
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59326841"
---
# <a name="manage-calculation-mode-in-excel"></a>Gérer le mode de calcul dans Excel

Cet exemple montre comment utiliser le [mode de calcul](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) et calculer des méthodes dans Excel sur le Web à l’aide Office scripts. Vous pouvez essayer le script sur n’importe Excel fichier.

## <a name="scenario"></a>Scénario

Le recalcul des workbooks avec un grand nombre de formules peut prendre un certain temps. Au lieu de laisser Excel contrôler le moment où les calculs ont lieu, vous pouvez les gérer dans le cadre de votre script. Cela permet d’améliorer les performances dans certains scénarios.

L’exemple de script définit le mode de calcul sur manuel. Cela signifie que le workbook ne recalcule les formules que lorsque le script l’indique (ou que vous calculez manuellement via [l’interface utilisateur).](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4) Le script affiche ensuite le mode de calcul actuel et recalcule entièrement le workbook entier.

## <a name="sample-code-control-calculation-mode"></a>Exemple de code : mode de calcul de contrôle

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

## <a name="training-video-manage-calculation-mode"></a>Vidéo de formation : gérer le mode de calcul

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/iw6O8QH01CI).
