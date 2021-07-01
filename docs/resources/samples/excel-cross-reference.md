---
title: Renvoi de fichiers Excel avec des Power Automate
description: Découvrez comment utiliser Office scripts et Power Automate pour faire référence à un fichier Excel format.
ms.date: 06/25/2021
localization_priority: Normal
ms.openlocfilehash: 89c4a5fa5dcff21681fa20cd4118447d39d9b6da
ms.sourcegitcommit: a063b3faf6c1b7c294bd6a73e46845b352f2a22d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/29/2021
ms.locfileid: "53202868"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="7a2ce-103">Renvoi de fichiers Excel avec des Power Automate</span><span class="sxs-lookup"><span data-stu-id="7a2ce-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="7a2ce-104">Cette solution montre comment comparer des données entre deux fichiers Excel pour rechercher des incohérences.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="7a2ce-105">Il utilise Office scripts pour analyser les données et Power Automate pour communiquer entre les workbooks.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="7a2ce-106">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="7a2ce-106">Example scenario</span></span>

<span data-ttu-id="7a2ce-107">Vous êtes un coordinateur d’événements qui est en train de planifier des haut-parleurs pour les conférences à venir.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="7a2ce-108">Vous conservez les données d’événement dans une feuille de calcul et les inscriptions du haut-parleur dans une autre.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="7a2ce-109">Pour vous assurer que les deux workbooks sont synchronisés, vous utilisez un flux avec Office scripts pour mettre en évidence les problèmes potentiels.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="7a2ce-110">Exemples Excel fichiers</span><span class="sxs-lookup"><span data-stu-id="7a2ce-110">Sample Excel files</span></span>

<span data-ttu-id="7a2ce-111">Téléchargez les fichiers suivants utilisés dans cette solution pour l’essayer vous-même !</span><span class="sxs-lookup"><span data-stu-id="7a2ce-111">Download the following files used in this solution to try it out yourself!</span></span>

1. <span data-ttu-id="7a2ce-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="7a2ce-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="7a2ce-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="7a2ce-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="7a2ce-114">Exemple de code : obtenir des données d’événement</span><span class="sxs-lookup"><span data-stu-id="7a2ce-114">Sample code: Get event data</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="7a2ce-115">Exemple de code : valider les inscriptions des haut-parleurs</span><span class="sxs-lookup"><span data-stu-id="7a2ce-115">Sample code: Validate speaker registrations</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
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
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="7a2ce-116">Power Automate flux de travail : vérifier les incohérences entre les workbooks</span><span class="sxs-lookup"><span data-stu-id="7a2ce-116">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="7a2ce-117">Ce flux extrait les informations d’événement du premier workbook et utilise ces données pour valider le second.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-117">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="7a2ce-118">Connectez-Power Automate et créez un flux **de cloud instantané.** [](https://flow.microsoft.com)</span><span class="sxs-lookup"><span data-stu-id="7a2ce-118">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="7a2ce-119">Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**</span><span class="sxs-lookup"><span data-stu-id="7a2ce-119">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="7a2ce-120">Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** avec l’action **de script Exécuter.**</span><span class="sxs-lookup"><span data-stu-id="7a2ce-120">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="7a2ce-121">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="7a2ce-121">Use the following values for the action:</span></span>
    * <span data-ttu-id="7a2ce-122">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="7a2ce-122">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="7a2ce-123">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="7a2ce-123">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="7a2ce-124">**Fichier**: event-data.xlsx ([sélectionné avec le sélecateur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="7a2ce-124">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="7a2ce-125">**Script**: obtenir des données d’événement</span><span class="sxs-lookup"><span data-stu-id="7a2ce-125">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Le connecteur Excel Online (Entreprise) pour le premier script dans Power Automate.":::

1. <span data-ttu-id="7a2ce-127">Ajoutez une deuxième **étape nouvelle** qui utilise le connecteur Excel **Online (Entreprise)** avec l’action **exécuter le script.**</span><span class="sxs-lookup"><span data-stu-id="7a2ce-127">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="7a2ce-128">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="7a2ce-128">Use the following values for the action:</span></span>
    * <span data-ttu-id="7a2ce-129">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="7a2ce-129">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="7a2ce-130">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="7a2ce-130">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="7a2ce-131">**Fichier**: speaker-registration.xlsx ([sélectionné avec le sélecateur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="7a2ce-131">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="7a2ce-132">**Script :** valider l’inscription du haut-parleur</span><span class="sxs-lookup"><span data-stu-id="7a2ce-132">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Le connecteur Excel Online (Entreprise) pour le deuxième script dans Power Automate.":::
1. <span data-ttu-id="7a2ce-134">Cet exemple utilise Outlook client de messagerie.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-134">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="7a2ce-135">Vous pouvez utiliser n’importe quel connecteur de messagerie Power Automate prend en charge.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-135">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="7a2ce-136">Ajoutez **une nouvelle étape** qui utilise le connecteur **Office 365 Outlook** et l’action Envoyer et e-mail **(V2).**</span><span class="sxs-lookup"><span data-stu-id="7a2ce-136">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="7a2ce-137">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="7a2ce-137">Use the following values for the action:</span></span>
    * <span data-ttu-id="7a2ce-138">**À**: Votre compte de messagerie de test (ou e-mail personnel)</span><span class="sxs-lookup"><span data-stu-id="7a2ce-138">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="7a2ce-139">**Objet**: Résultats de validation d’événement</span><span class="sxs-lookup"><span data-stu-id="7a2ce-139">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="7a2ce-140">**Body**: result (_dynamic content from Run script **2**_)</span><span class="sxs-lookup"><span data-stu-id="7a2ce-140">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Connecteur de Office 365 Outlook terminé dans Power Automate.":::
1. <span data-ttu-id="7a2ce-142">Enregistrez le flux, puis **sélectionnez Tester** pour l’essayer. Vous devriez recevoir un e-mail vous disant « Insérez une insérialisation trouvée.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-142">Save the flow, then select **Test** to try it out. You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="7a2ce-143">Les données nécessitent votre révision. »</span><span class="sxs-lookup"><span data-stu-id="7a2ce-143">Data requires your review."</span></span> <span data-ttu-id="7a2ce-144">Cela indique qu’il existe des différences entre les lignes dans **speaker-registrations.xlsx** et les lignes **dansevent-data.xlsx**.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-144">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="7a2ce-145">Ouvrez **speaker-registrations.xlsx** pour voir plusieurs cellules mises en surbrillation, où il existe des problèmes potentiels avec les listes d’inscription du haut-parleur.</span><span class="sxs-lookup"><span data-stu-id="7a2ce-145">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
