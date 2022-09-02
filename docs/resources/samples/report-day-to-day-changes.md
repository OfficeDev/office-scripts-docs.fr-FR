---
title: Enregistrer les modifications quotidiennes dans Excel et les signaler à l’aide d’un flux Power Automate
description: Découvrez comment utiliser les scripts Office et Power Automate pour suivre les modifications de valeur dans un classeur
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 083ca08573db060aa4788aea58fc67e50d004a4b
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572651"
---
# <a name="record-day-to-day-changes-in-excel-and-report-them-with-a-power-automate-flow"></a>Enregistrer les modifications quotidiennes dans Excel et les signaler à l’aide d’un flux Power Automate

Power Automate et les scripts Office se combinent pour gérer les tâches répétitives pour vous. Dans cet exemple, vous êtes chargé d’enregistrer une seule lecture numérique dans un classeur tous les jours et de signaler la modification depuis hier. Vous allez créer un flux pour obtenir cette lecture, le journaliser dans le classeur et signaler la modification par e-mail.

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez [daily-readings.xlsx](daily-readings.xlsx) pour un classeur prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-record-and-report-daily-readings"></a>Exemple de code : Enregistrer et signaler des lectures quotidiennes

```TypeScript
function main(workbook: ExcelScript.Workbook, newData: string): string {
  // Get the table by its name.
  const table = workbook.getTable("ReadingTable");

  // Read the current last entry in the Reading column.
  const readingColumn = table.getColumnByName("Reading");
  const readingColumnValues = readingColumn.getRange().getValues();
  const previousValue = readingColumnValues[readingColumnValues.length - 1][0] as number;

  // Add a row with the date, new value, and a formula calculating the difference.
  const currentDate = new Date(Date.now()).toLocaleDateString();
  const newRow = [currentDate, newData, "=[@Reading]-OFFSET([@Reading],-1,0)"];
  table.addRow(-1, newRow,);

  // Return the difference between the newData and the previous entry.
  const difference = Number.parseFloat(newData) - previousValue;
  console.log(difference);
  return difference;
}
```

## <a name="sample-flow-report-day-to-day-changes"></a>Exemple de flux : Signaler les modifications quotidiennes

Suivez ces étapes pour créer un flux [Power Automate](https://powerautomate.microsoft.com/) pour l’exemple.

1. Créez un **flux de cloud planifié**.
1. Planifiez la répétition du flux tous les **1 jour**.

    :::image type="content" source="../../images/day-to-day-changes-flow-1.png" alt-text="Étape de création du flux qui montre qu’elle se répète tous les jours.":::
1. Sélectionnez **Créer**.
1. Dans un flux réel, vous allez ajouter une étape qui obtient vos données. Les données peuvent provenir d’un autre classeur, d’une carte adaptative Teams ou de toute autre source. Pour tester l’exemple, effectuez un numéro de test. Ajoutez une nouvelle étape avec l’action **Initialiser la variable** . Donnez-lui les valeurs suivantes.
    1. **Nom** : Entrée
    1. **Type** : Entier
    1. **Valeur** : 190000

    :::image type="content" source="../../images/day-to-day-changes-flow-2.png" alt-text="Action d’initialisation de la variable avec les valeurs données.":::
1. Ajoutez une nouvelle étape avec le connecteur **Excel Online (Entreprise)** avec l’action **Exécuter le script** . Utilisez les valeurs suivantes pour l’action.
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **Fichier** : daily-readings.xlsx *(choisi par le biais du navigateur de fichiers)*
    1. **Script** : nom de votre script
    1. **newData** : Entrée *(contenu dynamique)*

    :::image type="content" source="../../images/day-to-day-changes-flow-3.png" alt-text="Action Exécuter le script avec les valeurs données.":::
1. Le script retourne la différence de lecture quotidienne sous forme de contenu dynamique nommé « result ». Pour l’exemple, vous pouvez envoyer les informations par e-mail à vous-même. Créez une étape qui utilise le connecteur **Outlook** avec l’action **Envoyer un e-mail (V2)** (ou le client de messagerie de votre choix). Utilisez les valeurs suivantes pour terminer l’action.
    1. **À** : votre adresse e-mail
    1. **Objet** : Modification de la lecture quotidienne
    1. **Corps** : résultat « Différence par rapport à hier » *(contenu dynamique d’Excel)*

    :::image type="content" source="../../images/day-to-day-changes-flow-4.png" alt-text="Connecteur Outlook terminé dans Power Automate.":::
1. Enregistrez le flux et essayez-le. Utilisez le bouton **Test** dans la page de l’éditeur de flux. Veillez à autoriser l’accès lorsque vous y êtes invité.
