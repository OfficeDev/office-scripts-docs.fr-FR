---
title: Filtrer le tableau Excel et obtenir une plage visible
description: Découvrez comment utiliser des scripts Office pour filtrer un tableau Excel et obtenir la plage visible sous la mesure d’un tableau d’objets.
ms.date: 03/16/2021
localization_priority: Normal
ms.openlocfilehash: c0a5842af4a62162225e3fc10203c261b91e010a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571241"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Filtrer le tableau Excel et obtenir une plage visible en tant qu’objet JSON

Cet exemple filtre un tableau Excel et renvoie la plage visible en tant qu’objet JSON. Ce JSON peut être fourni à un flux Power Automate dans le cadre d’une solution plus grande.

## <a name="example-scenario"></a>Exemple de scénario

* Appliquer un filtre à une colonne de tableau.
* Extraire la plage visible après le filtrage.
* Assemblez et renvoyer un objet avec une [structure JSON spécifique.](#sample-json)

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Exemple de code : filtrer un tableau et obtenir une plage visible

Le script suivant filtre un tableau et obtient la plage visible.

Téléchargez l’exemple <a href="table-filter.xlsx">table-filter.xlsx</a> fichier et utilisez-le avec ce script pour l’essayer vous-même !

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
  const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);

  console.log(uniqueKeys);
  const returnObj: ReturnTemplate = {}

  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    const rangeView = table1.getRange().getVisibleView();
    returnObj[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  })
  table1.getColumnByName('Station').getFilter().clear();
  console.log(JSON.stringify(returnObj));
  return returnObj
}

function returnObjectFromValues(values: string[][]): BasicObj[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i=0; i < values.length; i++) {
    if (i===0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j=0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray;
}

interface BasicObj {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObj[]
}
```

### <a name="sample-json"></a>Exemple de JSON

Chaque clé représente une valeur unique d’un tableau. Chaque instance de tableau représente la ligne visible lorsque le filtre correspondant est appliqué.

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason": ""
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason": ""
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason": ""
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason": ""
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Vidéo de formation : filtrer un tableau Excel et obtenir la plage visible

[![Regardez une vidéo pas à pas sur la façon de filtrer un tableau Excel et d’obtenir la plage visible](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Vidéo pas à pas sur la façon de filtrer un tableau Excel et d’obtenir la plage visible")
