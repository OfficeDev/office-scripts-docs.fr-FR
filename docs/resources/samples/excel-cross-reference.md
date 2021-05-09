---
title: Renvoi et mise en forme d’Excel fichier
description: Découvrez comment utiliser Office scripts et Power Automate pour faire référence à un fichier Excel format.
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: 7cc10787190e7ba8f5984ddda8b3c770eb0f7d8a
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285905"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="7d7b9-103">Renvoi et mise en forme d’Excel fichier</span><span class="sxs-lookup"><span data-stu-id="7d7b9-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="7d7b9-104">Cette solution montre comment deux fichiers Excel peuvent être référencés et formatés à l’aide Office scripts et Power Automate.</span><span class="sxs-lookup"><span data-stu-id="7d7b9-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="7d7b9-105">Le projet atteint les objectifs suivants :</span><span class="sxs-lookup"><span data-stu-id="7d7b9-105">The project achieves the following:</span></span>

1. <span data-ttu-id="7d7b9-106">Extrait les données d’événements <a href="events.xlsx">events.xlsx</a> l’aide d’une action de script Exécuter.</span><span class="sxs-lookup"><span data-stu-id="7d7b9-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="7d7b9-107">Transmet ces données au deuxième fichier Excel contenant les données de transaction d’événement et utilise ces données pour valider de base les données et mettre en forme des données manquantes ou incorrectes à l’aide de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="7d7b9-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="7d7b9-108">Envoie le résultat par courrier électronique à un réviseur.</span><span class="sxs-lookup"><span data-stu-id="7d7b9-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="7d7b9-109">Pour plus d’informations, voir Référence croisée et mise en forme de deux fichiers Excel à l’aide [Office scripts.](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)</span><span class="sxs-lookup"><span data-stu-id="7d7b9-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="7d7b9-110">Exemples Excel fichiers</span><span class="sxs-lookup"><span data-stu-id="7d7b9-110">Sample Excel files</span></span>

<span data-ttu-id="7d7b9-111">Téléchargez les fichiers suivants utilisés dans cette solution pour l’essayer vous-même !</span><span class="sxs-lookup"><span data-stu-id="7d7b9-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="7d7b9-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="7d7b9-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="7d7b9-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="7d7b9-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="7d7b9-114">Exemple de code : obtenir des données d’événement</span><span class="sxs-lookup"><span data-stu-id="7d7b9-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];
  
  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
      let [event, date, location, capacity] = row;
      records.push({
          event: event as string,
          date: date as number, 
          location: location as string,
          capacity: capacity as number
      })
  }

  // Log the event data to the console and return it for a flow.
  console.log(JSON.stringify(records));
  return records;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="7d7b9-115">Exemple de code : valider les transactions d’événement</span><span class="sxs-lookup"><span data-stu-id="7d7b9-115">Sample code: Validate event transactions</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);
    
 // Apply some basic formatting for readability.
  table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
  table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let overallMatch = true;

  // Iterate over every row.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [event, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyObject of keysObject) {
      if (keyObject.event === event) {
        match = true;

        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (keyObject.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (keyObject.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (keyObject.capacity !== capacity) {
          overallMatch = false;
          range.getCell(i, 3).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");      
    }  
  }

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  }
  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="7d7b9-116">Vidéo de formation : référence croisée et mise en forme d’un Excel de formation</span><span class="sxs-lookup"><span data-stu-id="7d7b9-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="7d7b9-117">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/dVwqBf483qo").</span><span class="sxs-lookup"><span data-stu-id="7d7b9-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
