---
title: Planifier des entretiens dans Teams
description: Découvrez comment utiliser les scripts Office pour envoyer une réunion Teams à partir de données Excel.
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e8c4af40398842e219dc3e2a80c6d2ee72d6b83
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572576"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Exemple de scénario Office Scripts : Planifier des entretiens dans Teams

Dans ce scénario, vous êtes un recruteur RH qui planifie des réunions d’entretien avec des candidats dans Teams. Vous gérez la planification des entretiens des candidats dans un fichier Excel. Vous devez envoyer l’invitation à la réunion Teams aux candidats et aux intervieweurs. Vous devez ensuite mettre à jour le fichier Excel avec la confirmation que des réunions Teams ont été envoyées.

La solution comporte trois étapes qui sont combinées dans un flux Power Automate unique.

1. Un script extrait des données d’une table et retourne un tableau d’objets sous forme de données [JSON](https://www.w3schools.com/whatis/whatis_json.asp) .
1. Les données sont ensuite envoyées à l’action **De créer une réunion Teams** pour envoyer des invitations.
1. Les mêmes données JSON sont envoyées à un autre script pour mettre à jour l’état de l’invitation.

Pour plus d’informations sur l’utilisation de JSON, consultez [Utiliser JSON pour transmettre des données vers et depuis des scripts Office](../../develop/use-json.md).

## <a name="scripting-skills-covered"></a>Compétences de script couvertes

* Flux Power Automate
* Intégration de Teams
* Analyse de table

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez le fichier [hr-schedule.xlsx](hr-schedule.xlsx) utilisé dans cette solution et essayez-le vous-même ! Veillez à modifier au moins une des adresses e-mail afin de recevoir une invitation.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Exemple de code : Extraire des données de table pour planifier des invitations

Ajoutez ce script à votre collection de scripts. Nommez-le **Planifier des entretiens** pour le flux.

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a>Exemple de code : Marquer les lignes comme invitées

Ajoutez ce script à votre collection de scripts. **Nommez-le Enregistrer les invites envoyées** pour le flux.

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Exemple de flux : exécuter les scripts de planification des entretiens et envoyer les réunions Teams

1. Créez un **flux de cloud instantané**.
1. Choisissez **déclencher manuellement un flux** , puis **sélectionnez Créer**.
1. Ajoutez une **nouvelle étape** qui utilise le connecteur **Excel Online (Entreprise)** et l’action **Exécuter le script** . Complétez le connecteur avec les valeurs suivantes.
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **Fichier** : hr-interviews.xlsx *(choisi par le biais du navigateur de fichiers)*
    1. **Script** : :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Capture d’écran planifier des entretiens du connecteur Excel Online (Entreprise) terminé pour obtenir des données d’entretien à partir du classeur dans Power Automate.":::
1. Ajoutez une **nouvelle étape** qui utilise l’action **Créer une réunion Teams** . Lorsque vous sélectionnez du contenu dynamique dans le connecteur Excel, une **application à chaque** bloc est générée pour votre flux. Complétez le connecteur avec les valeurs suivantes.
    1. **ID de calendrier** : Calendrier
    1. **Sujet**: Contoso Interview
    1. **Message** : **Message** (valeur Excel)
    1. **Fuseau horaire** : Heure standard du Pacifique
    1. **Heure de début** : **StartTime** (valeur Excel)
    1. **Heure de fin** : **FinishTime** (valeur Excel)
    1. **Participants obligatoires** : **CandidateEmail** ; **InterviewerEmail** (valeurs Excel) :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Capture d’écran du connecteur Teams terminé pour planifier des réunions dans Power Automate.":::
1. Dans la même **application à chaque** bloc, ajoutez un autre connecteur **Excel Online (Entreprise)** avec l’action **Exécuter le script** . Utilisez les valeurs ci-après.
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **Fichier** : hr-interviews.xlsx *(choisi par le biais du navigateur de fichiers)*
    1. **Script** : Enregistrer les invitations envoyées
    1. **invites** : **résultat** (valeur Excel) :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Capture d’écran du connecteur Excel Online (Entreprise) terminé pour enregistrer les invitations envoyées dans Power Automate.":::
1. Enregistrez le flux et essayez-le. Utilisez le bouton **Tester** dans la page de l’éditeur de flux ou exécutez le flux dans l’onglet **Mes flux** . Veillez à autoriser l’accès lorsque vous y êtes invité.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Vidéo de formation : Envoyer une réunion Teams à partir de données Excel

[Regardez Sudhi Ramamurthy parcourir une version de cet exemple sur YouTube](https://youtu.be/HyBdx52NOE8). Sa version utilise un script plus robuste qui gère l’évolution des colonnes et les heures de réunion obsolètes.
