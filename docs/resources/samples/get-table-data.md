---
title: Sortie Excel données en tant que JSON
description: Découvrez comment Excel données de table en tant que JSON à utiliser dans Power Automate.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: fefeda4f7e60880758f8f01e03f437a70c4111d4
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074570"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="f9dfb-103">Sortie Excel données de table en tant que JSON pour une utilisation Power Automate</span><span class="sxs-lookup"><span data-stu-id="f9dfb-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="f9dfb-104">Excel données de table peuvent être représentées sous la forme d’un tableau d’objets sous la forme de JSON.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="f9dfb-105">Chaque objet représente une ligne dans le tableau.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-105">Each object represents a row in the table.</span></span> <span data-ttu-id="f9dfb-106">Cela permet d’extraire les données Excel dans un format cohérent visible par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="f9dfb-107">Les données peuvent ensuite être données à d’autres systèmes via Power Automate flux.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="f9dfb-108">_Données de table d’entrée_</span><span class="sxs-lookup"><span data-stu-id="f9dfb-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="Feuille de calcul montrant les données de table d’entrée.":::

<span data-ttu-id="f9dfb-110">Une variante de cet exemple inclut également les liens hypertexte dans l’une des colonnes du tableau.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="f9dfb-111">Cela permet d’surfacer des niveaux supplémentaires de données de cellule dans le JSON.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="f9dfb-112">_Données de table d’entrée qui incluent des liens hypertexte_</span><span class="sxs-lookup"><span data-stu-id="f9dfb-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="Feuille de calcul montrant une colonne de données de tableau mise en forme sous forme de liens hypertexte.":::

<span data-ttu-id="f9dfb-114">_Boîte de dialogue pour modifier le lien hypertexte_</span><span class="sxs-lookup"><span data-stu-id="f9dfb-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="Boîte de dialogue Modifier le lien hypertexte affichant les options de modification des liens hypertexte.":::

## <a name="sample-excel-file"></a><span data-ttu-id="f9dfb-116">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="f9dfb-116">Sample Excel file</span></span>

<span data-ttu-id="f9dfb-117">Téléchargez le fichier <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> utilisé dans ces exemples et testez-le vous-même !</span><span class="sxs-lookup"><span data-stu-id="f9dfb-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="f9dfb-118">Exemple de code : renvoyer des données de table en tant que JSON</span><span class="sxs-lookup"><span data-stu-id="f9dfb-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="f9dfb-119">Vous pouvez modifier la `interface TableData` structure de façon à ce qu’elle corresponde à vos colonnes de tableau.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="f9dfb-120">Notez que pour les noms de colonnes avec des espaces, n’oubliez pas de placer votre clé entre guillemets, comme dans `"Event ID"` l’exemple.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('PlainTable').getTables()[0];

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

// This function converts a 2D-array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
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

  return objectArray as TableData[];
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output-from-the-plaintable-worksheet"></a><span data-ttu-id="f9dfb-121">Exemple de sortie de la feuille de calcul « PlainTable »</span><span class="sxs-lookup"><span data-stu-id="f9dfb-121">Sample output from the "PlainTable" worksheet</span></span>

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

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="f9dfb-122">Exemple de code : renvoyer des données de table en tant que JSON avec du texte de lien hypertexte</span><span class="sxs-lookup"><span data-stu-id="f9dfb-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="f9dfb-123">Le script extrait toujours les liens hypertexte de la quatrième colonne (0 index) de la table.</span><span class="sxs-lookup"><span data-stu-id="f9dfb-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="f9dfb-124">Vous pouvez modifier cet ordre ou inclure plusieurs colonnes en tant que données de lien hypertexte en modifiant le code sous le commentaire `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="f9dfb-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "WithHyperLink" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];

  // Get all the values from the table as text.
  const range = table.getRange();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(range);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(range: ExcelScript.Range): TableData[] {
  let values = range.getTexts();
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        object[objectKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        object[objectKeys[j]] = values[i][j];
      }
    }

    objectArray.push(object);
  }
  return objectArray as TableData[];
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

### <a name="sample-output-from-the-withhyperlink-worksheet"></a><span data-ttu-id="f9dfb-125">Exemple de sortie de la feuille de calcul « WithHyperLink »</span><span class="sxs-lookup"><span data-stu-id="f9dfb-125">Sample output from the "WithHyperLink" worksheet</span></span>

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

## <a name="use-in-power-automate"></a><span data-ttu-id="f9dfb-126">À utiliser dans Power Automate</span><span class="sxs-lookup"><span data-stu-id="f9dfb-126">Use in Power Automate</span></span>

<span data-ttu-id="f9dfb-127">Pour savoir comment utiliser un tel script dans Power Automate, voir Créer un flux de travail automatisé [avec Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span><span class="sxs-lookup"><span data-stu-id="f9dfb-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
