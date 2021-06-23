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
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a><span data-ttu-id="41dcc-103">Office Exemple de scénario de scripts : planifier des entretiens dans Teams</span><span class="sxs-lookup"><span data-stu-id="41dcc-103">Office Scripts sample scenario: Schedule interviews in Teams</span></span>

<span data-ttu-id="41dcc-104">Dans ce scénario, vous êtes un recrutement RH qui planifiera des réunions d’entretien avec des candidats Teams.</span><span class="sxs-lookup"><span data-stu-id="41dcc-104">In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams.</span></span> <span data-ttu-id="41dcc-105">Vous gérez le planning d’entretien des candidats dans Excel fichier.</span><span class="sxs-lookup"><span data-stu-id="41dcc-105">You manage the interview schedule of candidates in an Excel file.</span></span> <span data-ttu-id="41dcc-106">Vous devez envoyer l’invitation Teams réunion au candidat et aux intervieweurs.</span><span class="sxs-lookup"><span data-stu-id="41dcc-106">You'll need to send the Teams meeting invite to both the candidate and interviewers.</span></span> <span data-ttu-id="41dcc-107">Vous devez ensuite mettre à jour le fichier Excel avec la confirmation que Teams réunions ont été envoyées.</span><span class="sxs-lookup"><span data-stu-id="41dcc-107">You then need to update the Excel file with the confirmation that Teams meetings have been sent.</span></span>

<span data-ttu-id="41dcc-108">La solution possède trois étapes qui sont combinées en une seule Power Automate flux.</span><span class="sxs-lookup"><span data-stu-id="41dcc-108">The solution has three steps that are combined in a single Power Automate flow.</span></span>

1. <span data-ttu-id="41dcc-109">Un script extrait les données d’une table et renvoie un tableau d’objets en tant que données JSON.</span><span class="sxs-lookup"><span data-stu-id="41dcc-109">A script extracts data from a table and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="41dcc-110">Les données sont ensuite envoyées au Teams **créer une** action Teams réunion pour envoyer des invitations.</span><span class="sxs-lookup"><span data-stu-id="41dcc-110">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span>
1. <span data-ttu-id="41dcc-111">Les mêmes données JSON sont envoyées à un autre script pour mettre à jour l’état de l’invitation.</span><span class="sxs-lookup"><span data-stu-id="41dcc-111">The same JSON data is sent to another script to update the status of the invitation.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="41dcc-112">Compétences d’écriture de scripts couvertes</span><span class="sxs-lookup"><span data-stu-id="41dcc-112">Scripting skills covered</span></span>

* <span data-ttu-id="41dcc-113">Power Automate flux</span><span class="sxs-lookup"><span data-stu-id="41dcc-113">Power Automate flows</span></span>
* <span data-ttu-id="41dcc-114">Teams’intégration</span><span class="sxs-lookup"><span data-stu-id="41dcc-114">Teams integration</span></span>
* <span data-ttu-id="41dcc-115">Table parsing</span><span class="sxs-lookup"><span data-stu-id="41dcc-115">Table parsing</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="41dcc-116">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="41dcc-116">Sample Excel file</span></span>

<span data-ttu-id="41dcc-117">Téléchargez le fichier <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> utilisé dans cette solution et testez-le vous-même !</span><span class="sxs-lookup"><span data-stu-id="41dcc-117">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span> <span data-ttu-id="41dcc-118">N’oubliez pas de modifier au moins l’une des adresses de messagerie afin de recevoir une invitation.</span><span class="sxs-lookup"><span data-stu-id="41dcc-118">Be sure to change at least one of the email addresses so that you receive an invite.</span></span>

## <a name="sample-code-extract-table-data-to-schedule-invites"></a><span data-ttu-id="41dcc-119">Exemple de code : extraire des données de table pour planifier des invitations</span><span class="sxs-lookup"><span data-stu-id="41dcc-119">Sample code: Extract table data to schedule invites</span></span>

<span data-ttu-id="41dcc-120">Nommez ce script **Schedule Interviews** pour le flux.</span><span class="sxs-lookup"><span data-stu-id="41dcc-120">Name this script **Schedule Interviews** for the flow.</span></span>

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

## <a name="sample-code-mark-rows-as-invited"></a><span data-ttu-id="41dcc-121">Exemple de code : marquer les lignes comme invitées</span><span class="sxs-lookup"><span data-stu-id="41dcc-121">Sample code: Mark rows as invited</span></span>

<span data-ttu-id="41dcc-122">Nommez ce script **Enregistrer les invitations envoyées** pour le flux.</span><span class="sxs-lookup"><span data-stu-id="41dcc-122">Name this script **Record Sent Invites** for the flow.</span></span>

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

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a><span data-ttu-id="41dcc-123">Exemple de flux : exécuter les scripts de planification d’entretien et envoyer Teams réunions</span><span class="sxs-lookup"><span data-stu-id="41dcc-123">Sample flow: Run the interview scheduling scripts and send the Teams meetings</span></span>

1. <span data-ttu-id="41dcc-124">Créez un **flux de cloud instantané.**</span><span class="sxs-lookup"><span data-stu-id="41dcc-124">Create a new **Instant cloud flow**.</span></span>
1. <span data-ttu-id="41dcc-125">Sélectionnez **Déclencher manuellement un flux et** appuyez sur **Créer.**</span><span class="sxs-lookup"><span data-stu-id="41dcc-125">Select **Manually trigger a flow** and press **Create**.</span></span>
1. <span data-ttu-id="41dcc-126">Ajoutez **une nouvelle étape** qui utilise le connecteur Excel Online **(Entreprise)** et l’action **de script Exécuter.**</span><span class="sxs-lookup"><span data-stu-id="41dcc-126">Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action.</span></span> <span data-ttu-id="41dcc-127">Terminez le connecteur avec les valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="41dcc-127">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="41dcc-128">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="41dcc-128">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="41dcc-129">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="41dcc-129">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="41dcc-130">**Fichier**: hr-interviews.xlsx *(choisi via le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="41dcc-130">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. **Script**: Planifier des entretiens Capture d’écran du connecteur :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Excel Online (Entreprise)"::: pour obtenir des données d’entretien à partir du Power Automate.
1. <span data-ttu-id="41dcc-132">Ajoutez **une nouvelle étape** qui utilise l’action Créer Teams **réunion.**</span><span class="sxs-lookup"><span data-stu-id="41dcc-132">Add a **New step** that uses the **Create a Teams meeting** action.</span></span> <span data-ttu-id="41dcc-133">Lorsque vous sélectionnez du contenu dynamique à partir du connecteur Excel, une application à chaque **bloc** est générée pour votre flux.</span><span class="sxs-lookup"><span data-stu-id="41dcc-133">As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow.</span></span> <span data-ttu-id="41dcc-134">Terminez le connecteur avec les valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="41dcc-134">Complete the connector with the following values.</span></span>
    1. <span data-ttu-id="41dcc-135">**ID de calendrier**: Calendrier</span><span class="sxs-lookup"><span data-stu-id="41dcc-135">**Calendar id**: Calendar</span></span>
    1. <span data-ttu-id="41dcc-136">**Objet**: Contoso Interview</span><span class="sxs-lookup"><span data-stu-id="41dcc-136">**Subject**: Contoso Interview</span></span>
    1. <span data-ttu-id="41dcc-137">**Message**: **Message** (valeur Excel)</span><span class="sxs-lookup"><span data-stu-id="41dcc-137">**Message**: **Message** (the Excel value)</span></span>
    1. <span data-ttu-id="41dcc-138">**Fuseau horaire**: heure standard du Pacifique</span><span class="sxs-lookup"><span data-stu-id="41dcc-138">**Time zone**: Pacific Standard Time</span></span>
    1. <span data-ttu-id="41dcc-139">**Heure de** début **: StartTime** (valeur Excel)</span><span class="sxs-lookup"><span data-stu-id="41dcc-139">**Start time**: **StartTime** (the Excel value)</span></span>
    1. <span data-ttu-id="41dcc-140">**Heure de fin** **: FinishTime** (valeur Excel)</span><span class="sxs-lookup"><span data-stu-id="41dcc-140">**End time**: **FinishTime** (the Excel value)</span></span>
    1. **Participants obligatoires**: **CandidateEmail** ; **ScreenshotEmail** (les valeurs Excel) Capture d’écran du connecteur Teams terminé pour planifier des réunions :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="dans Power Automate.":::
1. <span data-ttu-id="41dcc-142">Dans le même **bloc, ajoutez** un autre connecteur **Excel Online (Entreprise)** avec l’action **exécuter le script.**</span><span class="sxs-lookup"><span data-stu-id="41dcc-142">In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action.</span></span> <span data-ttu-id="41dcc-143">Utilisez les valeurs ci-après.</span><span class="sxs-lookup"><span data-stu-id="41dcc-143">Use the following values.</span></span>
    1. <span data-ttu-id="41dcc-144">**Emplacement** : OneDrive Entreprise</span><span class="sxs-lookup"><span data-stu-id="41dcc-144">**Location**: OneDrive for Business</span></span>
    1. <span data-ttu-id="41dcc-145">**Bibliothèque de documents** : OneDrive</span><span class="sxs-lookup"><span data-stu-id="41dcc-145">**Document Library**: OneDrive</span></span>
    1. <span data-ttu-id="41dcc-146">**Fichier**: hr-interviews.xlsx *(choisi via le navigateur de fichiers)*</span><span class="sxs-lookup"><span data-stu-id="41dcc-146">**File**: hr-interviews.xlsx *(Chosen through the file browser)*</span></span>
    1. <span data-ttu-id="41dcc-147">**Script**: enregistrer les invitations envoyées</span><span class="sxs-lookup"><span data-stu-id="41dcc-147">**Script**: Record Sent Invites</span></span>
    1. **invite :** **résultat (valeur** Excel) Capture d’écran du connecteur Excel Online (Entreprise) terminé pour enregistrer que les invitations ont été :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="envoyées dans Power Automate.":::
1. <span data-ttu-id="41dcc-149">Enregistrez le flux et testez-le.</span><span class="sxs-lookup"><span data-stu-id="41dcc-149">Save the flow and try it out.</span></span>

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="41dcc-150">Vidéo de formation : envoyer une Teams réunion à partir Excel données</span><span class="sxs-lookup"><span data-stu-id="41dcc-150">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="41dcc-151">[Regardez Sudhi Genrethy parcourir une version de cet exemple sur YouTube](https://youtu.be/HyBdx52NOE8).</span><span class="sxs-lookup"><span data-stu-id="41dcc-151">[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span> <span data-ttu-id="41dcc-152">Sa version utilise un script plus robuste qui gère la modification des colonnes et des heures de réunion obsolètes.</span><span class="sxs-lookup"><span data-stu-id="41dcc-152">His version uses a more robust script that handles changing columns and obsolete meeting times.</span></span>
