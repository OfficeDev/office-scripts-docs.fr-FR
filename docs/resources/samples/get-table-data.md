---
title: Sortie de données Excel en tant que JSON
description: Découvrez comment créer des données de tableau Excel en tant que JSON à utiliser dans Power Automate.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 678506fee0b6a41ede8245fb360d485d635e2d64
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571216"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="f0f3b-103">Sortie des données de tableau Excel en tant que JSON pour une utilisation dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="f0f3b-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="f0f3b-104">Les données de tableau Excel peuvent être représentées sous la forme d’un tableau d’objets sous la forme de JSON.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="f0f3b-105">Chaque objet représente une ligne dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-105">Each object represents a row in the table.</span></span> <span data-ttu-id="f0f3b-106">Cela permet d’extraire les données d’Excel dans un format cohérent visible par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="f0f3b-107">Les données peuvent ensuite être données à d’autres systèmes via des flux Power Automate.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="f0f3b-108">_Données de table d’entrée_</span><span class="sxs-lookup"><span data-stu-id="f0f3b-108">_Input table data_</span></span>

![Capture d’écran affichant les données de la table d’entrée](../../images/table-input.png)

<span data-ttu-id="f0f3b-110">Une variante de cet exemple inclut également les liens hypertexte dans l’une des colonnes du tableau.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="f0f3b-111">Cela permet d’surfacer des niveaux supplémentaires de données de cellule dans le JSON.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="f0f3b-112">_Données de table d’entrée incluant des liens hypertexte_</span><span class="sxs-lookup"><span data-stu-id="f0f3b-112">_Input table data that includes hyperlinks_</span></span>

![Capture d’écran affichant des données de tableau qui incluent des liens hypertexte](../../images/table-hyperlink-view.png)

<span data-ttu-id="f0f3b-114">_Boîte de dialogue pour modifier le lien hypertexte_</span><span class="sxs-lookup"><span data-stu-id="f0f3b-114">_Dialog to edit hyperlink_</span></span>

![Capture d’écran affichant une boîte de dialogue pour modifier le lien hypertexte](../../images/table-hyperlink-edit.png)

## <a name="sample-excel-file"></a><span data-ttu-id="f0f3b-116">Exemple de fichier Excel</span><span class="sxs-lookup"><span data-stu-id="f0f3b-116">Sample Excel file</span></span>

<span data-ttu-id="f0f3b-117">Téléchargez le fichier <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> utilisés dans ces exemples et testez-le vous-même !</span><span class="sxs-lookup"><span data-stu-id="f0f3b-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="f0f3b-118">Exemple de code : renvoyer des données de table en tant que JSON</span><span class="sxs-lookup"><span data-stu-id="f0f3b-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="f0f3b-119">Vous pouvez modifier la `interface TableData` structure de façon à ce qu’elle corresponde à vos colonnes de tableau.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="f0f3b-120">Notez que pour les noms de colonnes avec des espaces, n’oubliez pas de placer votre clé entre guillemets, comme dans `"Event ID"` l’exemple.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('PlainTable').getTables()[0];
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][]): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output"></a><span data-ttu-id="f0f3b-121">Sortie d’exemple</span><span class="sxs-lookup"><span data-stu-id="f0f3b-121">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="f0f3b-122">Exemple de code : renvoyer des données de table en tant que JSON avec du texte de lien hypertexte</span><span class="sxs-lookup"><span data-stu-id="f0f3b-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="f0f3b-123">Le script extrait toujours les liens hypertexte de la quatrième colonne (0 index) de la table.</span><span class="sxs-lookup"><span data-stu-id="f0f3b-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="f0f3b-124">Vous pouvez modifier cet ordre ou inclure plusieurs colonnes en tant que données de lien hypertexte en modifiant le code sous le commentaire `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="f0f3b-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];
  const range = table.getRange();
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts, range);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][], range: ExcelScript.Range): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        obj[objKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        obj[objKeys[j]] = values[i][j];
      }
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  "Search link": string
  Speakers: string
}
```

### <a name="sample-output"></a><span data-ttu-id="f0f3b-125">Sortie d’exemple</span><span class="sxs-lookup"><span data-stu-id="f0f3b-125">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="f0f3b-126">Utilisation dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="f0f3b-126">Use in Power Automate</span></span>

<span data-ttu-id="f0f3b-127">Pour savoir comment utiliser un tel script dans Power Automate, voir Créer un flux de travail [automatisé avec Power Automate.](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="f0f3b-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
