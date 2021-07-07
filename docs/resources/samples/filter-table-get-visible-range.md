---
title: Filtrer Excel tableau et obtenir une plage visible
description: Découvrez comment utiliser des scripts Office pour filtrer un tableau Excel et obtenir la plage visible en tant que tableau d’objets.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: b19b826f95c7e7aeb331130fde05afaafe500c3d
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313952"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="dee57-103">Filtrer Excel tableau et obtenir une plage visible en tant qu’objet JSON</span><span class="sxs-lookup"><span data-stu-id="dee57-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="dee57-104">Cet exemple filtre un tableau Excel et renvoie la plage visible en tant qu’objet JSON.</span><span class="sxs-lookup"><span data-stu-id="dee57-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="dee57-105">Ce JSON peut être fourni à un flux Power Automate dans le cadre d’une solution plus grande.</span><span class="sxs-lookup"><span data-stu-id="dee57-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="dee57-106">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="dee57-106">Example scenario</span></span>

* <span data-ttu-id="dee57-107">Appliquer un filtre à une colonne de tableau.</span><span class="sxs-lookup"><span data-stu-id="dee57-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="dee57-108">Extraire la plage visible après le filtrage.</span><span class="sxs-lookup"><span data-stu-id="dee57-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="dee57-109">Assemblez et renvoyer un objet avec une [structure JSON spécifique.](#sample-json)</span><span class="sxs-lookup"><span data-stu-id="dee57-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="dee57-110">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="dee57-110">Sample Excel file</span></span>

<span data-ttu-id="dee57-111">Téléchargez <a href="table-filter.xlsx">table-filter.xlsx</a> pour un livre de travail prêt à l’emploi.</span><span class="sxs-lookup"><span data-stu-id="dee57-111">Download <a href="table-filter.xlsx">table-filter.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="dee57-112">Ajoutez le script suivant pour essayer l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="dee57-112">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="dee57-113">Exemple de code : filtrer un tableau et obtenir une plage visible</span><span class="sxs-lookup"><span data-stu-id="dee57-113">Sample code: Filter a table and get visible range</span></span>

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
  let objectArray = [];
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

### <a name="sample-json"></a><span data-ttu-id="dee57-114">Exemple de JSON</span><span class="sxs-lookup"><span data-stu-id="dee57-114">Sample JSON</span></span>

<span data-ttu-id="dee57-115">Chaque clé représente une valeur unique d’une table.</span><span class="sxs-lookup"><span data-stu-id="dee57-115">Each key represents a unique value of a table.</span></span> <span data-ttu-id="dee57-116">Chaque instance de tableau représente la ligne visible lorsque le filtre correspondant est appliqué.</span><span class="sxs-lookup"><span data-stu-id="dee57-116">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="dee57-117">Vidéo de formation : filtrer un tableau Excel et obtenir la plage visible</span><span class="sxs-lookup"><span data-stu-id="dee57-117">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="dee57-118">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/Mv7BrvPq84A).</span><span class="sxs-lookup"><span data-stu-id="dee57-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/Mv7BrvPq84A).</span></span>
