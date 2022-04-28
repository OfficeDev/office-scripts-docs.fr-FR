---
title: 'Office scénario d’exemple de scripts : bouton d’horloge punch'
description: Cet exemple ajoute un bouton d’horloge perforé et permet à un utilisateur d’entrer et d’expirer à l’aide de l’heure actuelle.
ms.date: 04/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: de56fb651d6f6088620678cfd72ce662875eafa7
ms.sourcegitcommit: e6428a5214fa38aef036a952a0e3c09dbf6e4d3e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/28/2022
ms.locfileid: "65109291"
---
# <a name="office-scripts-sample-scenario-punch-clock-button"></a>Office scénario d’exemple de scripts : bouton d’horloge punch

L’idée de scénario et le script utilisés dans cet exemple ont été contribués par Office membre de la communauté [Scripts Brian Gonzalez](https://github.com/b-gonzalez).

Dans ce scénario, vous allez créer une feuille de temps pour un employé qui lui permet d’enregistrer ses heures de début et de fin en appuyant sur un [bouton](../../develop/script-buttons.md). En fonction de ce qui a été enregistré précédemment, appuyer sur le bouton démarre sa journée (horloge) ou met fin à sa journée (expiration de l’horloge). L’exemple fonctionne à la fois pour Excel sur le Web et sur Windows.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="Tableau avec trois colonnes (« Clock In », « Clock Out » et « Duration ») et un bouton intitulé « Punch clock » dans le classeur.":::

## <a name="setup-instructions"></a>Instructions d’installation

1. Téléchargez <a href="punch-clock-sample.xlsx">punch-clock-sample.xlsx</a> sur votre OneDrive.

    :::image type="content" source="../../images/punch-clock-sample-1.png" alt-text="Tableau avec trois colonnes : « Clock In », « Clock Out » et « Duration ».":::

1. Ouvrez le classeur dans Excel sur le Web.

1. Sous l’onglet **Automatiser** , sélectionnez **Nouveau script** et collez le script suivant dans l’éditeur.

    ```typescript
    /**
     * This script records either the start or end time of a shift, 
     * depending on what is filled out in the table. 
     * It is intended to be used with a Script Button.
     */
    function main(workbook: ExcelScript.Workbook) {
      // Get the first table in the timesheet.
      const timeSheet = workbook.getWorksheet("MyTimeSheet");
      const timeTable = timeSheet.getTables()[0];
    
      // Get the appropriate table columns.
      const clockInColumn = timeTable.getColumnByName("Clock In");
      const clockOutColumn = timeTable.getColumnByName("Clock Out");
      const durationColumn = timeTable.getColumnByName("Duration");
    
      // Get the last rows for the Clock In and Clock Out columns.
      let clockInLastRow = clockInColumn.getRangeBetweenHeaderAndTotal().getLastRow();
      let clockOutLastRow = clockOutColumn.getRangeBetweenHeaderAndTotal().getLastRow();
    
      // Get the current date to use as the start or end time.
      let date: Date = new Date();
    
      // Add the current time to a column based on the state of the table.
      if (clockInLastRow.getValue() as string === "") {
        // If the Clock In column has an empty value in the table, add a start time.
        clockInLastRow.setValue(date.toLocaleString());
      } else if (clockOutLastRow.getValue() as string === "") {
        // If the Clock Out column has an empty value in the table, 
        // add an end time and calculate the shift duration.
        clockOutLastRow.setValue(date.toLocaleString());
        const clockInTime = new Date(clockInLastRow.getValue() as string);
        const clockOutTime  = new Date(clockOutLastRow.getValue() as string);
        const clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()));
    
        let durationString = getDurationMessage(clockDuration);
        durationColumn.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
      } else {
        // If both columns are full, add a new row, then add a start time.
        timeTable.addRow()
        clockInLastRow.getOffsetRange(1, 0).setValue(date.toLocaleString());
      }
    }
    
    /**
     * A function to write a time duration as a string.
     */
    function getDurationMessage(delta: number) {
      // Adapted from here:
      // https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    
      delta = delta / 1000;
      let durationString = "";
    
      let days = Math.floor(delta / 86400);
      delta -= days * 86400;
    
      let hours = Math.floor(delta / 3600) % 24;
      delta -= hours * 3600;
    
      let minutes = Math.floor(delta / 60) % 60;
    
      if (days >= 1) {
        durationString += days;
        durationString += (days > 1 ? " days" : " day");
    
        if (hours >= 1 && minutes >= 1) {
          durationString += ", ";
        }
        else if (hours >= 1 || minutes > 1) {
          durationString += " and ";
        }
      }
    
      if (hours >= 1) {
        durationString += hours;
        durationString += (hours > 1 ? " hours" : " hour");
        if (minutes >= 1) {
          durationString += " and ";
        }
      }
    
      if (minutes >= 1) {
        durationString += minutes;
        durationString += (minutes > 1 ? " minutes" : " minute");
      }
    
      return durationString;
    }
    ```

1. Renommez le script « Punch clock ».

1. Enregistrez le script.

1. Dans le classeur, sélectionnez la cellule **E2**.

1. Ajouter un bouton de script. Accédez au menu **Plus d’options (...)** dans la page **Détails du script** , puis sélectionnez **Bouton Ajouter**.

    :::image type="content" source="../../images/punch-clock-sample-2.png" alt-text="Menu « Autres options » et bouton « Bouton Ajouter ».":::

1. Enregistrez le classeur.

## <a name="run-the-script"></a>Exécuter le script

Appuyez sur le bouton **d’horloge Punch** pour exécuter le script. Il enregistre l’heure actuelle sous « Clock In » ou « Clock Out », en fonction de ce qui a été précédemment entré.

:::image type="content" source="../../images/punch-clock-sample-3.png" alt-text="La table et le bouton « Punch clock » dans le classeur.":::

> [!NOTE]
> La durée n’est enregistrée que si elle est supérieure à une minute. Modifiez manuellement l’heure « Horloge » pour tester des durées plus longues.
