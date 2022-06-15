---
title: Filtrer Excel table et obtenir une plage visible
description: Découvrez comment utiliser Office Scripts pour filtrer une table Excel et obtenir la plage visible sous la forme d’un tableau d’objets.
ms.date: 03/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 103ec97111720ab872c0be843aa0573781d98c44
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088084"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>Filtrer Excel table et obtenir une plage visible en tant qu’objet JSON

Cet exemple filtre une table Excel et retourne la plage visible sous la forme d’un objet [JSON](https://www.w3schools.com/whatis/whatis_json.asp). Ce JSON peut être fourni à un flux de Power Automate dans le cadre d’une solution plus grande.

## <a name="example-scenario"></a>Exemple de scénario

* Appliquez un filtre à une colonne de table.
* Extrayez la plage visible après le filtrage.
* Assemblez et retournez un objet avec une [structure JSON spécifique](#sample-json).

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez <a href="table-filter.xlsx">table-filter.xlsx</a> pour un classeur prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>Exemple de code : Filtrer une table et obtenir une plage visible

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  // Get the "Station" column to use as key values in the filter.
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(value => value[0] as string);

  // Filter out repeated keys. This call to `filter` only returns the first instance of every unique element in the array.
  const uniqueKeys = keyColumnValues.filter((value, index, array) => array.indexOf(value) === index);
  console.log(uniqueKeys);

  const stationData: ReturnTemplate = {};

  // Filter the table to show only rows corresponding to each key.
  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    
    // Get the visible view when a single filter is active.
    const rangeView = table1.getRange().getVisibleView();

    // Create a JSON object with every visible row.
    stationData[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  });

  // Remove the filters.
  table1.getColumnByName('Station').getFilter().clear();

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(stationData));
  return stationData;
}

// This function converts a 2D-array of values into a generic JSON object.
function returnObjectFromValues(values: string[][]): BasicObject[] {
  let objectArray: BasicObject[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object);
  }

  return objectArray;
}

interface BasicObject {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObject[]
}
```

### <a name="sample-json"></a>Exemple de code JSON

Chaque clé représente une valeur unique d’une table. Chaque instance de tableau représente la ligne visible lorsque le filtre correspondant est appliqué. Pour plus d’informations sur l’utilisation de JSON, consultez [Utiliser JSON pour transmettre des données vers et depuis Office Scripts](../../develop/use-json.md).

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>Vidéo de formation : Filtrer une table Excel et obtenir la plage visible

[Regardez Sudhi Ramamurthy parcourir cet exemple sur YouTube](https://youtu.be/Mv7BrvPq84A).
