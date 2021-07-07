---
title: Renvoi de fichiers Excel avec des Power Automate
description: Découvrez comment utiliser Office scripts et Power Automate pour faire référence à un fichier Excel format.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 0776ce49cacecfa15339cc7c0cd4866daad789ff
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313959"
---
# <a name="cross-reference-excel-files-with-power-automate"></a><span data-ttu-id="37b2d-103">Renvoi de fichiers Excel avec des Power Automate</span><span class="sxs-lookup"><span data-stu-id="37b2d-103">Cross-reference Excel files with Power Automate</span></span>

<span data-ttu-id="37b2d-104">Cette solution montre comment comparer des données entre deux fichiers Excel pour rechercher des incohérences.</span><span class="sxs-lookup"><span data-stu-id="37b2d-104">This solution shows how to compare data across two Excel files to find discrepancies.</span></span> <span data-ttu-id="37b2d-105">Il utilise Office scripts pour analyser les données et Power Automate pour communiquer entre les workbooks.</span><span class="sxs-lookup"><span data-stu-id="37b2d-105">It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="37b2d-106">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="37b2d-106">Example scenario</span></span>

<span data-ttu-id="37b2d-107">Vous êtes un coordinateur d’événements qui est en train de planifier des haut-parleurs pour les conférences à venir.</span><span class="sxs-lookup"><span data-stu-id="37b2d-107">You're an event coordinator who is scheduling speakers for upcoming conferences.</span></span> <span data-ttu-id="37b2d-108">Vous conservez les données d’événement dans une feuille de calcul et les inscriptions du haut-parleur dans une autre.</span><span class="sxs-lookup"><span data-stu-id="37b2d-108">You keep the event data in one spreadsheet and the speaker registrations in another.</span></span> <span data-ttu-id="37b2d-109">Pour vous assurer que les deux workbooks sont synchronisés, vous utilisez un flux avec Office scripts pour mettre en évidence les problèmes potentiels.</span><span class="sxs-lookup"><span data-stu-id="37b2d-109">To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.</span></span>

## <a name="sample-excel-files"></a><span data-ttu-id="37b2d-110">Exemples Excel fichiers</span><span class="sxs-lookup"><span data-stu-id="37b2d-110">Sample Excel files</span></span>

<span data-ttu-id="37b2d-111">Téléchargez les fichiers suivants pour obtenir des workbooks prêts à l’emploi pour l’exemple.</span><span class="sxs-lookup"><span data-stu-id="37b2d-111">Download the following files to get ready-to-use workbooks for the sample.</span></span>

1. <span data-ttu-id="37b2d-112"><a href="event-data.xlsx">event-data.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="37b2d-112"><a href="event-data.xlsx">event-data.xlsx</a></span></span>
1. <span data-ttu-id="37b2d-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span><span class="sxs-lookup"><span data-stu-id="37b2d-113"><a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a></span></span>

<span data-ttu-id="37b2d-114">Ajoutez les scripts suivants pour essayer l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="37b2d-114">Add the following scripts to try the sample yourself!</span></span>

## <a name="sample-code-get-event-data"></a><span data-ttu-id="37b2d-115">Exemple de code : obtenir des données d’événement</span><span class="sxs-lookup"><span data-stu-id="37b2d-115">Sample code: Get event data</span></span>

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

## <a name="sample-code-validate-speaker-registrations"></a><span data-ttu-id="37b2d-116">Exemple de code : valider les inscriptions des haut-parleurs</span><span class="sxs-lookup"><span data-stu-id="37b2d-116">Sample code: Validate speaker registrations</span></span>

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a><span data-ttu-id="37b2d-117">Power Automate flux de travail : vérifier les incohérences entre les workbooks</span><span class="sxs-lookup"><span data-stu-id="37b2d-117">Power Automate flow: Check for inconsistencies across the workbooks</span></span>

<span data-ttu-id="37b2d-118">Ce flux extrait les informations d’événement du premier workbook et utilise ces données pour valider le second.</span><span class="sxs-lookup"><span data-stu-id="37b2d-118">This flow extracts the event information from the first workbook and uses that data to validate the second workbook.</span></span>

1. <span data-ttu-id="37b2d-119">Connectez-Power Automate et créez un flux **de cloud instantané.** [](https://flow.microsoft.com)</span><span class="sxs-lookup"><span data-stu-id="37b2d-119">Sign into [Power Automate](https://flow.microsoft.com) and create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="37b2d-120">Sélectionnez **Déclencher manuellement un flux,** puis **sélectionnez Créer.**</span><span class="sxs-lookup"><span data-stu-id="37b2d-120">Choose **Manually trigger a flow** and select **Create**.</span></span>
1. <span data-ttu-id="37b2d-121">Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** avec l’action **de script Exécuter.**</span><span class="sxs-lookup"><span data-stu-id="37b2d-121">Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="37b2d-122">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="37b2d-122">Use the following values for the action:</span></span>
    * <span data-ttu-id="37b2d-123">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="37b2d-123">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="37b2d-124">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="37b2d-124">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="37b2d-125">**Fichier**: event-data.xlsx ([sélectionné avec le sélecateur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="37b2d-125">**File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="37b2d-126">**Script**: obtenir des données d’événement</span><span class="sxs-lookup"><span data-stu-id="37b2d-126">**Script**: Get event data</span></span>

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Le connecteur Excel Online (Entreprise) pour le premier script dans Power Automate.":::

1. <span data-ttu-id="37b2d-128">Ajoutez une deuxième **étape nouvelle** qui utilise le connecteur Excel **Online (Entreprise)** avec l’action **exécuter le script.**</span><span class="sxs-lookup"><span data-stu-id="37b2d-128">Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="37b2d-129">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="37b2d-129">Use the following values for the action:</span></span>
    * <span data-ttu-id="37b2d-130">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="37b2d-130">**Location**: OneDrive for Business</span></span>
    * <span data-ttu-id="37b2d-131">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="37b2d-131">**Document Library**: OneDrive</span></span>
    * <span data-ttu-id="37b2d-132">**Fichier**: speaker-registration.xlsx ([sélectionné avec le sélecateur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span><span class="sxs-lookup"><span data-stu-id="37b2d-132">**File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))</span></span>
    * <span data-ttu-id="37b2d-133">**Script :** valider l’inscription du haut-parleur</span><span class="sxs-lookup"><span data-stu-id="37b2d-133">**Script**: Validate speaker registration</span></span>

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Le connecteur Excel Online (Entreprise) pour le deuxième script dans Power Automate.":::
1. <span data-ttu-id="37b2d-135">Cet exemple utilise Outlook client de messagerie.</span><span class="sxs-lookup"><span data-stu-id="37b2d-135">This sample uses Outlook as the email client.</span></span> <span data-ttu-id="37b2d-136">Vous pouvez utiliser n’importe quel connecteur de messagerie Power Automate prend en charge.</span><span class="sxs-lookup"><span data-stu-id="37b2d-136">You could use any email connector Power Automate supports.</span></span> <span data-ttu-id="37b2d-137">Ajoutez **une nouvelle étape** qui utilise le connecteur **Office 365 Outlook** et l’action Envoyer et e-mail **(V2).**</span><span class="sxs-lookup"><span data-stu-id="37b2d-137">Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action.</span></span> <span data-ttu-id="37b2d-138">Utilisez les valeurs suivantes pour l’action :</span><span class="sxs-lookup"><span data-stu-id="37b2d-138">Use the following values for the action:</span></span>
    * <span data-ttu-id="37b2d-139">**À**: Votre compte de messagerie de test (ou e-mail personnel)</span><span class="sxs-lookup"><span data-stu-id="37b2d-139">**To**: Your test email account (or personal email)</span></span>
    * <span data-ttu-id="37b2d-140">**Objet**: Résultats de validation d’événement</span><span class="sxs-lookup"><span data-stu-id="37b2d-140">**Subject**: Event validation results</span></span>
    * <span data-ttu-id="37b2d-141">**Body**: result (_dynamic content from Run script **2**_)</span><span class="sxs-lookup"><span data-stu-id="37b2d-141">**Body**: result (_dynamic content from **Run script 2**_)</span></span>

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Connecteur de Office 365 Outlook terminé dans Power Automate.":::
1. <span data-ttu-id="37b2d-143">Enregistrez le flux.</span><span class="sxs-lookup"><span data-stu-id="37b2d-143">Save the flow.</span></span> <span data-ttu-id="37b2d-144">Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.</span><span class="sxs-lookup"><span data-stu-id="37b2d-144">Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.</span></span>
1. <span data-ttu-id="37b2d-145">Vous devriez recevoir un e-mail vous disant « Insérez une insérialisation trouvée.</span><span class="sxs-lookup"><span data-stu-id="37b2d-145">You should receive an email saying "Mismatch found.</span></span> <span data-ttu-id="37b2d-146">Les données nécessitent votre révision. »</span><span class="sxs-lookup"><span data-stu-id="37b2d-146">Data requires your review."</span></span> <span data-ttu-id="37b2d-147">Cela indique qu’il existe des différences entre les lignes dans **speaker-registrations.xlsx** et les lignes **dansevent-data.xlsx**.</span><span class="sxs-lookup"><span data-stu-id="37b2d-147">This indicates there are differences between rows in **speaker-registrations.xlsx** and rows in **event-data.xlsx**.</span></span> <span data-ttu-id="37b2d-148">Ouvrez **speaker-registrations.xlsx** pour voir plusieurs cellules mises en surbrillation, où il existe des problèmes potentiels avec les listes d’inscription du haut-parleur.</span><span class="sxs-lookup"><span data-stu-id="37b2d-148">Open **speaker-registrations.xlsx** to see several highlighted cells where there are potential problems with the speaker registration listings.</span></span>
