---
title: Filtrer Excel tableau et obtenir une plage visible
description: Découvrez comment utiliser des scripts Office pour filtrer un tableau Excel et obtenir la plage visible en tant que tableau d’objets.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a310857e6055b3da57c353dc7ad78a6fbdd86d4e
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232374"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Filtrer Excel tableau et obtenir une plage visible en tant qu’objet JSON

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
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason&quot;: &quot;"
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason&quot;: &quot;"
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason&quot;: &quot;"
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Vidéo de formation : filtrer un tableau Excel et obtenir la plage visible

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/Mv7BrvPq84A).
