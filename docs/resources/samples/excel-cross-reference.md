---
title: Renvoi de fichiers Excel avec des Power Automate
description: Découvrez comment utiliser Office scripts et Power Automate pour faire référence à un fichier Excel format.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: adeb84140cb9884309c9f37854a29fc4d59b17ed
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332977"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Renvoi de fichiers Excel avec des Power Automate

Cette solution montre comment comparer des données entre deux fichiers Excel pour rechercher des incohérences. Il utilise Office scripts pour analyser les données et Power Automate pour communiquer entre les workbooks.

## <a name="example-scenario"></a>Exemple de scénario

Vous êtes un coordinateur d’événements qui est en train de planifier des haut-parleurs pour les conférences à venir. Vous conservez les données d’événement dans une feuille de calcul et les inscriptions du haut-parleur dans une autre. Pour vous assurer que les deux workbooks sont synchronisés, vous utilisez un flux avec Office scripts pour mettre en évidence les problèmes potentiels.

## <a name="sample-excel-files"></a>Exemples Excel fichiers

Téléchargez les fichiers suivants pour obtenir des workbooks prêts à l’emploi pour l’exemple.

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

Ajoutez les scripts suivants pour essayer l’exemple vous-même !

## <a name="sample-code-get-event-data"></a>Exemple de code : obtenir des données d’événement

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

## <a name="sample-code-validate-speaker-registrations"></a>Exemple de code : valider les inscriptions des haut-parleurs

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Power Automate flux de travail : vérifier les incohérences entre les workbooks

Ce flux extrait les informations d’événement du premier workbook et utilise ces données pour valider le second.

1. Connectez-Power Automate et créez un flux **de cloud instantané.** [](https://flow.microsoft.com)
1. Sélectionnez **Déclencher manuellement un flux,** puis **sélectionnez Créer.**
1. Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** avec l’action **de script Exécuter.** Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier**: event-data.xlsx ([sélectionné avec le sélecateur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: obtenir des données d’événement

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Le connecteur Excel Online (Entreprise) pour le premier script dans Power Automate.":::

1. Ajoutez une deuxième **étape nouvelle** qui utilise le connecteur Excel **Online (Entreprise)** avec l’action **exécuter le script.** Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier**: speaker-registration.xlsx ([sélectionné avec le sélecateur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script :** valider l’inscription du haut-parleur

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Le connecteur Excel Online (Entreprise) pour le deuxième script dans Power Automate.":::
1. Cet exemple utilise Outlook client de messagerie. Vous pouvez utiliser n’importe quel connecteur de messagerie Power Automate prend en charge. Ajoutez **une nouvelle étape** qui utilise le connecteur **Office 365 Outlook** et l’action Envoyer et e-mail **(V2).** Utilisez les valeurs suivantes pour l’action.
    * **À**: Votre compte de messagerie de test (ou e-mail personnel)
    * **Objet**: Résultats de validation d’événement
    * **Body**: result (_dynamic content from Run script **2**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="Connecteur de Office 365 Outlook terminé dans Power Automate.":::
1. Enregistrez le flux. Utilisez le **bouton Test** sur la page de l’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.
1. Vous devriez recevoir un e-mail vous disant « Insérez une insérialisation trouvée. Les données nécessitent votre révision. » Cela indique qu’il existe des différences entre les lignes dans **speaker-registrations.xlsx** et les lignes **dansevent-data.xlsx**. Ouvrez **speaker-registrations.xlsx** pour voir plusieurs cellules mises en surbrillation, où il existe des problèmes potentiels avec les listes d’inscription du haut-parleur.
