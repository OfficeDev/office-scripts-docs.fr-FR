---
title: Envoyer une réunion Teams à partir Excel données
description: Découvrez comment utiliser les scripts Office pour envoyer une réunion de Teams à partir Excel données.
ms.date: 05/06/2021
localization_priority: Normal
ROBOTS: NOINDEX
ms.openlocfilehash: 85b39d7e3d1008dee01e7fe9c690116be1d7e5d8
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545629"
---
# <a name="send-teams-meeting-from-excel-data"></a><span data-ttu-id="079b9-103">Envoyer Teams réunion à partir Excel données</span><span class="sxs-lookup"><span data-stu-id="079b9-103">Send Teams meeting from Excel data</span></span>

<span data-ttu-id="079b9-104">Cette solution montre comment utiliser les scripts Office et les actions de Power Automate pour sélectionner des lignes à partir d’Excel un fichier et l’utiliser pour envoyer une invitation à une réunion de Teams puis mettre à jour Excel.</span><span class="sxs-lookup"><span data-stu-id="079b9-104">This solution shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send a Teams meeting invite then update Excel.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="079b9-105">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="079b9-105">Example scenario</span></span>

* <span data-ttu-id="079b9-106">Un recruteur rh gère le calendrier d’entrevue des candidats dans un Excel dossier.</span><span class="sxs-lookup"><span data-stu-id="079b9-106">An HR recruiter manages the interview schedule of candidates in an Excel file.</span></span>
* <span data-ttu-id="079b9-107">Le recruteur doit envoyer le Teams réunion inviter le candidat et les intervieweurs.</span><span class="sxs-lookup"><span data-stu-id="079b9-107">The recruiter needs to send the Teams meeting invite to the candidate and interviewers.</span></span> <span data-ttu-id="079b9-108">Les règles d’entreprise sont de sélectionner :</span><span class="sxs-lookup"><span data-stu-id="079b9-108">The business rules are to select:</span></span>

    <span data-ttu-id="079b9-109">a) Invite uniquement à ceux pour qui l’invitation n’est pas déjà envoyée comme enregistré dans la colonne de fichiers.</span><span class="sxs-lookup"><span data-stu-id="079b9-109">(a) Invites to only those for whom the invite isn't already sent as recorded in the file column.</span></span>

    <span data-ttu-id="079b9-110">b) Dates d’entrevue à l’avenir (pas de dates antérieures).</span><span class="sxs-lookup"><span data-stu-id="079b9-110">(b) Interview dates in the future (no past dates).</span></span>

* <span data-ttu-id="079b9-111">Le recruteur doit mettre à jour le dossier Excel avec la confirmation que toutes les réunions Teams ont été envoyées pour les dossiers admissibles.</span><span class="sxs-lookup"><span data-stu-id="079b9-111">The recruiter needs to update the Excel file with the confirmation that all Teams meetings have been sent for the eligible records.</span></span>

<span data-ttu-id="079b9-112">La solution a 3 parties:</span><span class="sxs-lookup"><span data-stu-id="079b9-112">The solution has 3 parts:</span></span>

1. <span data-ttu-id="079b9-113">Office Script pour extraire des données d’une table en fonction des conditions et renvoie un éventail d’objets sous forme de données JSON.</span><span class="sxs-lookup"><span data-stu-id="079b9-113">Office Script to extract data from a table based on conditions and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="079b9-114">Les données sont ensuite envoyées à l’Teams **créer une action Teams réunion** pour envoyer des invitations.</span><span class="sxs-lookup"><span data-stu-id="079b9-114">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span> <span data-ttu-id="079b9-115">Envoyez une Teams par instance dans le tableau JSON.</span><span class="sxs-lookup"><span data-stu-id="079b9-115">Send one Teams meeting per instance in the JSON array.</span></span>
1. <span data-ttu-id="079b9-116">Envoyez les mêmes données JSON à un autre script Office pour mettre à jour l’état de l’invitation.</span><span class="sxs-lookup"><span data-stu-id="079b9-116">Send the same JSON data to another Office Script to update the status of the invitation.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="079b9-117">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="079b9-117">Sample Excel file</span></span>

<span data-ttu-id="079b9-118">Téléchargez le <a href="hr-schedule.xlsx"> fichierhr-schedule.xlsxutilisé </a> dans cette solution et essayez-le vous-même!</span><span class="sxs-lookup"><span data-stu-id="079b9-118">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span>

## <a name="sample-code-select-filtered-rows-from-table-as-json"></a><span data-ttu-id="079b9-119">Exemple de code : Sélectionnez les lignes filtrées à partir de la table comme JSON</span><span class="sxs-lookup"><span data-stu-id="079b9-119">Sample code: Select filtered rows from table as JSON</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  console.log("Current date time: " + new Date().toUTCString());
  const MEETING_DURATION = workbook.getNamedItem('MeetingDuration').getRange().getValue() as number;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet('Interviews');
  const table = sheet.getTables()[0];
  const dataRows: string[][] = table.getRangeBetweenHeaderAndTotal().getTexts();

  // Convert the table rows into InterviewInvite objects for the flow.
  const recordDetails: RecordDetail[] = returnObjectFromValues(dataRows);
  const inviteRecords = generateInterviewRecords(recordDetails, MEETING_DURATION);
  console.log(JSON.stringify(inviteRecords));
  return inviteRecords;
}

/**
 * Converts table values into a RecordDetail array.
 */
function returnObjectFromValues(values: string[][]): RecordDetail[] {
  let objectArray: BasicObj[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }
    objectArray.push(object);
  }
  return objectArray as RecordDetail[];
}

/**
 * Generate interview records by selecting required columns.
 * @param records Input records from the table of interviews.
 * @param mins Number of minutes to add to the start date-time.
 */
function generateInterviewRecords(records: RecordDetail[], mins: number): InterviewInvite[] {
  const interviewInvites: InterviewInvite[] = [];

  records.forEach((record) => {
    // Interviewer 1
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time1'])) > new Date()) {
      console.log("selected " + new Date(record['Start time1']).toUTCString());
      let startTime = new Date(record['Start time1']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time1']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer1,
        InterviewerEmail: record['Interviewer1 email'],
        StartTime: startTime,
        FinishTime: finishTime
      });
    } else {
      console.log("Rejected " + (new Date(record['Start time1']).toUTCString()));
    }
    // Interviewer 2 
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time2'])) > new Date()) {
      console.log("selected " + new Date(record['Start time2']).toUTCString());


      let startTime = new Date(record['Start time2']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time2']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer2,
        InterviewerEmail: record['Interviewer2 email'],
        StartTime: startTime,
        FinishTime: finishTime
      })
    } else {
      console.log("Rejected " + (new Date(record['Start time2']).toUTCString()))

    }
  })
  return interviewInvites;
}

/**
 * Add minutes to start date-time.
 * @param startDateTime Start date-time
 * @param mins Minutes to add to the start date-time
 */
function addMins(startDateTime: Date, mins: number) {
  return new Date(startDateTime.getTime() + mins * 60 * 1000);
}

// Basic key-value pair object.
interface BasicObj {
  [key: string]: string | number | boolean
}

// Input record that matches the table data.
interface RecordDetail extends BasicObj {
  ID: string
  'Invite to interview': string
  Candidate: string
  'Candidate email': string
  'Candidate contact': string
  Interviewer1: string
  'Interviewer1 email': string
  Interviewer2: string
  'Interviewer2 email': string
  'Start time1': string
  'Start time2': string
}

// Output record.
interface InterviewInvite extends BasicObj {
  ID: string
  Candidate: string
  CandidateEmail: string
  CandidateContact: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
}
```

## <a name="sample-code-mark-as-invited"></a><span data-ttu-id="079b9-120">Exemple de code: Marquer comme invité</span><span class="sxs-lookup"><span data-stu-id="079b9-120">Sample code: Mark as invited</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, completedInvitesString: string) {
    completedInvitesString = `[
      {
        "ID": "10",
        "Candidate": "Adele ",
        "CandidateEmail": "AdeleV@M365x904181.OnMicrosoft.com",
        "CandidateContact": "1234567899",
        "Interviewer": "Megan",
        "InterviewerEmail": "MeganB@M365x904181.OnMicrosoft.com",
        "StartTime": "2020-11-03T18:30:00Z",
        "FinishTime": "2020-11-03T22:45:00Z"
      },
      {
        "ID": "30",
        "Candidate": "Allan ",
        "CandidateEmail": "AllanD@M365x904181.OnMicrosoft.com",
        "CandidateContact": "1234567978",
        "Interviewer": "Raul",
        "InterviewerEmail": "RaulR@M365x904181.OnMicrosoft.com",
        "StartTime": "2020-11-03T23:00:00Z",
        "FinishTime": "2020-11-03T23:45:00Z"
      }
    ]`;
    let completedInvites = JSON.parse(completedInvitesString) as InterviewInvite[];
    const sheet = workbook.getWorksheet('Interviews');
    const range = sheet.getTables()[0].getRange();
    const dataRows = range.getValues();
    for (let i=0; i < dataRows.length; i++) {
        for (let invite of completedInvites) {
            if (String(dataRows[i][0]) === invite.ID) {
                range.getCell(i,1).setValue(true);
            }
        }
    }
    return;
}


// Invite record.
interface InterviewInvite  {
    ID: string
    Candidate: string
    CandidateEmail: string
    CandidateContact: string
    Interviewer: string
    InterviewerEmail: string
    StartTime: string
    FinishTime: string
}
```

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="079b9-121">Vidéo de formation : Envoyer une réunion Teams à partir Excel données</span><span class="sxs-lookup"><span data-stu-id="079b9-121">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="079b9-122">[Regardez Sudhi Ramamurthy marcher à travers cet échantillon sur YouTube](https://youtu.be/HyBdx52NOE8).</span><span class="sxs-lookup"><span data-stu-id="079b9-122">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span>
