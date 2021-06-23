---
title: Planifier des entretiens dans Teams
description: Découvrez comment utiliser des scripts Office pour envoyer une Teams à partir de Excel données.
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: 66dae536c4a51ff3e028f06bf3aef3c7509d83bb
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074430"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office Exemple de scénario de scripts : planifier des entretiens dans Teams

Dans ce scénario, vous êtes un recrutement RH qui planifiera des réunions d’entretien avec des candidats Teams. Vous gérez le planning d’entretien des candidats dans Excel fichier. Vous devez envoyer l’invitation Teams réunion au candidat et aux intervieweurs. Vous devez ensuite mettre à jour le fichier Excel avec la confirmation que Teams réunions ont été envoyées.

La solution possède trois étapes qui sont combinées en une seule Power Automate flux.

1. Un script extrait les données d’une table et renvoie un tableau d’objets en tant que données JSON.
1. Les données sont ensuite envoyées au Teams **créer une** action Teams réunion pour envoyer des invitations.
1. Les mêmes données JSON sont envoyées à un autre script pour mettre à jour l’état de l’invitation.

## <a name="scripting-skills-covered"></a>Compétences d’écriture de scripts couvertes

* Power Automate flux
* Teams’intégration
* Table parsing

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez le fichier <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> utilisé dans cette solution et testez-le vous-même ! N’oubliez pas de modifier au moins l’une des adresses de messagerie afin de recevoir une invitation.

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>Exemple de code : extraire des données de table pour planifier des invitations

Nommez ce script **Schedule Interviews** pour le flux.

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

## <a name="sample-code-mark-rows-as-invited"></a>Exemple de code : marquer les lignes comme invitées

Nommez ce script **Enregistrer les invitations envoyées** pour le flux.

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>Exemple de flux : exécuter les scripts de planification d’entretien et envoyer Teams réunions

1. Créez un **flux de cloud instantané.**
1. Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**
1. Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** et l’action **de script Exécuter.** Terminez le connecteur avec les valeurs suivantes.
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **Fichier**: hr-interviews.xlsx *(choisi via le navigateur de fichiers)*
    1. **Script**: Planifier des entretiens Capture d’écran du connecteur :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Excel Online (Entreprise)"::: pour obtenir des données d’entretien à partir du Power Automate.
1. Ajoutez **une nouvelle étape** qui utilise l’action Créer Teams **réunion.** Lorsque vous sélectionnez du contenu dynamique à partir du connecteur Excel, une application à chaque **bloc** est générée pour votre flux. Terminez le connecteur avec les valeurs suivantes.
    1. **ID de calendrier**: Calendrier
    1. **Objet**: Contoso Interview
    1. **Message**: **Message** (valeur Excel)
    1. **Fuseau horaire**: heure standard du Pacifique
    1. **Heure de** début **: StartTime** (valeur Excel)
    1. **Heure de fin** **: FinishTime** (valeur Excel)
    1. **Participants obligatoires**: **CandidateEmail** ; **ScreenshotEmail** (les valeurs Excel) Capture d’écran du connecteur Teams terminé pour planifier des réunions :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="dans Power Automate.":::
1. Dans le même **bloc, ajoutez** un autre connecteur **Excel Online (Entreprise)** avec l’action **exécuter le script.** Utilisez les valeurs ci-après.
    1. **Emplacement** : OneDrive Entreprise
    1. **Bibliothèque de documents** : OneDrive
    1. **Fichier**: hr-interviews.xlsx *(choisi via le navigateur de fichiers)*
    1. **Script**: enregistrer les invitations envoyées
    1. **invite :** **résultat (valeur** Excel) Capture d’écran du connecteur Excel Online (Entreprise) terminé pour enregistrer que les invitations ont été :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="envoyées dans Power Automate.":::
1. Enregistrez le flux et testez-le.

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>Vidéo de formation : envoyer une Teams réunion à partir Excel données

[Regardez Sudhi Genrethy parcourir une version de cet exemple sur YouTube](https://youtu.be/HyBdx52NOE8). Sa version utilise un script plus robuste qui gère la modification des colonnes et des heures de réunion obsolètes.
