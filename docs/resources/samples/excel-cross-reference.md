---
title: Références croisées de fichiers Excel avec Power Automate
description: Découvrez comment utiliser les scripts Office et Power Automate pour référencer et mettre en forme un fichier Excel.
ms.date: 06/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b32249dc7cb1e8c1b841a4db6caaff3b4d2998ec
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572674"
---
# <a name="cross-reference-excel-files-with-power-automate"></a>Références croisées de fichiers Excel avec Power Automate

Cette solution montre comment comparer des données entre deux fichiers Excel pour rechercher des différences. Il utilise les scripts Office pour analyser les données et Power Automate pour communiquer entre les classeurs.

Cet exemple transmet des données entre des classeurs à l’aide d’objets [JSON](https://www.w3schools.com/whatis/whatis_json.asp) . Pour plus d’informations sur l’utilisation de JSON, consultez [Utiliser JSON pour transmettre des données vers et depuis des scripts Office](../../develop/use-json.md).

## <a name="example-scenario"></a>Exemple de scénario

Vous êtes un coordinateur d’événements qui planifie des conférenciers pour les prochaines conférences. Vous conservez les données d’événement dans une feuille de calcul et les inscriptions de l’orateur dans une autre. Pour vous assurer que les deux classeurs sont synchronisés, vous utilisez un flux avec les scripts Office pour mettre en évidence les éventuels problèmes.

## <a name="sample-excel-files"></a>Exemples de fichiers Excel

Téléchargez les fichiers suivants pour obtenir des classeurs prêts à l’emploi pour l’exemple.

1. [event-data.xlsx](event-data.xlsx)
1. [speaker-registrations.xlsx](speaker-registrations.xlsx)

Ajoutez les scripts suivants pour essayer l’exemple vous-même !

## <a name="sample-code-get-event-data"></a>Exemple de code : Obtenir des données d’événement

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

## <a name="sample-code-validate-speaker-registrations"></a>Exemple de code : Valider les inscriptions d’orateurs

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

## <a name="power-automate-flow-check-for-inconsistencies-across-the-workbooks"></a>Flux Power Automate : Rechercher les incohérences entre les classeurs

Ce flux extrait les informations d’événement du premier classeur et utilise ces données pour valider le deuxième classeur.

1. Connectez-vous à [Power Automate](https://flow.microsoft.com) et créez un **flux de cloud instantané**.
1. Choisissez **déclencher manuellement un flux** , puis **sélectionnez Créer**.
1. Ajoutez une **nouvelle étape** qui utilise le connecteur **Excel Online (Entreprise)** avec l’action **Exécuter le script** . Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier** : event-data.xlsx ([sélectionné avec le sélecteur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script** : obtenir des données d’événement

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="Connecteur Excel Online (Entreprise) terminé pour le premier script dans Power Automate.":::

1. Ajoutez une deuxième **étape nouvelle** qui utilise le connecteur **Excel Online (Entreprise)** avec l’action **Exécuter le script** . Cela utilise les valeurs retournées à partir du script **obtenir des données d’événement** comme entrée pour le script **valider les données d’événement** . Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier** : speaker-registration.xlsx ([sélectionné avec le sélecteur de fichiers](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script** : Valider l’inscription de l’orateur
    * **keys**: result (_dynamic content from **Run script**_)

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="Connecteur Excel Online (Entreprise) terminé pour le deuxième script dans Power Automate.":::
1. Cet exemple utilise Outlook comme client de messagerie. Vous pouvez utiliser n’importe quel connecteur de messagerie pris en charge par Power Automate. Ajoutez une **nouvelle étape** qui utilise le **connecteur Office 365 Outlook** et l’action **Envoyer et envoyer un e-mail (V2**). Cela utilise les valeurs retournées par le script **Valider l’inscription de l’orateur** comme contenu du corps de l’e-mail. Utilisez les valeurs suivantes pour l’action.
    * **À** : votre compte de messagerie de test (ou e-mail personnel)
    * **Objet** : Résultats de validation d’événement
    * **Corps** : résultat (_contenu dynamique du **script d’exécution 2**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="L’Office 365 connecteur Outlook terminé dans Power Automate.":::
1. Enregistrez le flux. Utilisez le bouton **Tester** dans la page de l’éditeur de flux ou exécutez le flux dans l’onglet **Mes flux** . Veillez à autoriser l’accès lorsque vous y êtes invité.
1. Vous devez recevoir un e-mail indiquant « Incompatibilité trouvée. Les données nécessitent votre révision. » Cela indique qu’il existe des différences entre les lignes dans **speaker-registrations.xlsx** et les lignes dans **event-data.xlsx**. Ouvrez **speaker-registrations.xlsx** pour voir plusieurs cellules mises en surbrillance où il existe des problèmes potentiels avec les listes d’inscription de l’orateur.
