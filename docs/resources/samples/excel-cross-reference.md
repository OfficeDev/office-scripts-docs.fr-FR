---
title: Renvoi et mise en forme d’un fichier Excel
description: Découvrez comment utiliser Les scripts Office et Power Automate pour faire référence à un fichier Excel et le mettre en forme.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 287de604733b7e6a126d0c81cb4e23351e558c61
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571272"
---
# <a name="cross-reference-and-format-an-excel-file"></a><span data-ttu-id="05f3e-103">Renvoi et mise en forme d’un fichier Excel</span><span class="sxs-lookup"><span data-stu-id="05f3e-103">Cross-reference and format an Excel file</span></span>

<span data-ttu-id="05f3e-104">Cette solution montre comment deux fichiers Excel peuvent être référencés et formatés à l’aide de Scripts Office et de Power Automate.</span><span class="sxs-lookup"><span data-stu-id="05f3e-104">This solution shows how two Excel files can be cross-referenced and formatted using Office Scripts and Power Automate.</span></span>

<span data-ttu-id="05f3e-105">Le projet atteint les objectifs suivants :</span><span class="sxs-lookup"><span data-stu-id="05f3e-105">The project achieves the following:</span></span>

1. <span data-ttu-id="05f3e-106">Extrait les données d’événements <a href="events.xlsx">events.xlsx</a> l’aide d’une action de script Exécuter.</span><span class="sxs-lookup"><span data-stu-id="05f3e-106">Extracts event data from <a href="events.xlsx">events.xlsx</a> using one Run script action.</span></span>
1. <span data-ttu-id="05f3e-107">Transmet ces données au deuxième fichier Excel contenant les données de transaction d’événements et utilise ces données pour valider de base les données et la mise en forme des données manquantes ou incorrectes à l’aide de Scripts Office.</span><span class="sxs-lookup"><span data-stu-id="05f3e-107">Passes that data to the second Excel file containing event transaction data and uses that data to do basic validation of data and formatting of missing or incorrect data using Office Scripts.</span></span>
1. <span data-ttu-id="05f3e-108">Envoie le résultat par courrier électronique à un réviseur.</span><span class="sxs-lookup"><span data-stu-id="05f3e-108">Emails the result to a reviewer.</span></span>

<span data-ttu-id="05f3e-109">Pour plus d’informations, voir Référence croisée et mise en forme de [deux fichiers Excel à l’aide de scripts Office.](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)</span><span class="sxs-lookup"><span data-stu-id="05f3e-109">For further details, see [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535).</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="05f3e-110">Exemples de fichiers Excel</span><span class="sxs-lookup"><span data-stu-id="05f3e-110">Sample Excel files</span></span>

<span data-ttu-id="05f3e-111">Téléchargez les fichiers suivants utilisés dans cette solution pour l’essayer vous-même !</span><span class="sxs-lookup"><span data-stu-id="05f3e-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="05f3e-112"><a href="events.xlsx">events.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="05f3e-112"><a href="events.xlsx">events.xlsx</a></span></span>
1. <span data-ttu-id="05f3e-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="05f3e-113"><a href="event-transactions.xlsx">event-transactions.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="05f3e-114">Exemple de code : obtenir des données d’événement</span><span class="sxs-lookup"><span data-stu-id="05f3e-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): EventData[] {
    let table = workbook.getWorksheet('Keys').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();
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
    console.log(JSON.stringify(records))
    return records;
}

interface EventData {
    event: string
    date: number
    location: string
    capacity: number
}
```

## <a name="sample-code-validate-event-transactions"></a><span data-ttu-id="05f3e-115">Exemple de code : valider les transactions d’événement</span><span class="sxs-lookup"><span data-stu-id="05f3e-115">Sample code: Validate event transactions</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
    let table = workbook.getWorksheet('Transactions').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    range.clear(ExcelScript.ClearApplyTo.formats);
  
    let overallMatch = true;
  
    table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
    table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
      .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    let rows = range.getValues();
    let keysObject = JSON.parse(keys) as EventData[];
    for (let i=0; i < rows.length; i++){
      let row = rows[i];
      let [event, date, location, capacity] = row;
      let match = false;
      for (let keyObject of keysObject){
        if (keyObject.event === event) {
          match = true;
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
      if (!match) {
        overallMatch = false;
        range.getCell(i, 0).getFormat()
          .getFill()
          .setColor("FFFF00");      
      }
  
    }
    let returnString = "All the data is in the right order.";
    if (overallMatch === false) {
      returnString = "Mismatch found. Data requires your review.";
    }
    console.log("Returning: " + returnString);
    return returnString;
}

interface EventData {
event: string
date: number
location: string
capacity: number
}
```

## <a name="training-video-cross-reference-and-format-an-excel-file"></a><span data-ttu-id="05f3e-116">Vidéo de formation : référence croisée et mise en forme d’un fichier Excel</span><span class="sxs-lookup"><span data-stu-id="05f3e-116">Training video: Cross-reference and format an Excel file</span></span>

<span data-ttu-id="05f3e-117">[![Regardez une vidéo pas à pas sur la façon de référencer et de mettre en forme un fichier Excel](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Vidéo pas à pas sur la façon de référencer et de mettre en forme un fichier Excel")</span><span class="sxs-lookup"><span data-stu-id="05f3e-117">[![Watch step-by-step video on how to cross-reference and format an Excel file](../../images/cross-ref-tables-vid.jpg)](https://youtu.be/dVwqBf483qo "Step-by-step video on how to cross-reference and format an Excel file")</span></span>
