---
title: Recoupement et format d’un fichier Excel’eau
description: Apprenez à utiliser les scripts Office et les Power Automate pour recouper et formater un fichier Excel texte.
ms.date: 05/06/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: f07395eb4e6c77b7aee3776e3252d135bc690a6f
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545765"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="9f699-103">Recoupement et format d’un fichier Excel’eau</span><span class="sxs-lookup"><span data-stu-id="9f699-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="9f699-104">Cette solution montre comment deux fichiers Excel peuvent être recoupés et formatés à l’aide de scripts Office scripts et de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="9f699-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="9f699-105">Le projet réalise les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="9f699-105">The project achieves the following:</span></span>

1. <span data-ttu-id="9f699-106">Extrait les données d’événements à <a href="events.xlsx"> partir deevents.xlsx'une </a> action de script Run.</span><span class="sxs-lookup"><span data-stu-id="9f699-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="9f699-107">Transmet ces données au deuxième fichier Excel contenant des données de transaction d’événements et utilise ces données pour valider de base les données et formater les données manquantes ou incorrectes à l’aide de scripts Office.</span><span class="sxs-lookup"><span data-stu-id="9f699-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="9f699-108">Envoie le résultat par courriel à un examinateur.</span><span class="sxs-lookup"><span data-stu-id="9f699-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="9f699-109">Pour plus de [détails, consultez Cross Reference et formater deux fichiers Excel à l’aide Office scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span><span class="sxs-lookup"><span data-stu-id="9f699-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="9f699-110">Exemples Excel fichiers</span><span class="sxs-lookup"><span data-stu-id="9f699-110">Sample Excel files</span></span>

<span data-ttu-id="9f699-111">Téléchargez les fichiers suivants utilisés dans cette solution pour l’essayer vous-même!</span><span class="sxs-lookup"><span data-stu-id="9f699-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="9f699-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="9f699-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="9f699-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="9f699-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="9f699-114">Exemple de code : Obtenez des données d’événement</span><span class="sxs-lookup"><span data-stu-id="9f699-114">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="9f699-115">Exemple de code : Valider les transactions événement</span><span class="sxs-lookup"><span data-stu-id="9f699-115">Sample code: Validate event transactions</span></span>

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

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="9f699-116">Vidéo de formation : Référence croisée et format d’un fichier Excel vidéo</span><span class="sxs-lookup"><span data-stu-id="9f699-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="9f699-117">[Regardez Sudhi Ramamurthy marcher à travers cet échantillon sur YouTube](https://youtu.be/dVwqBf483qo").</span><span class="sxs-lookup"><span data-stu-id="9f699-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/dVwqBf483qo").</span></span>
