---
title: Ajouter des commentaires dans Excel
description: Découvrez comment utiliser Office scripts pour ajouter des commentaires dans une feuille de calcul.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 77e308d020281c71751e2652f8dbaec00c263e44
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313910"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="d69d7-103">Ajouter des commentaires dans Excel</span><span class="sxs-lookup"><span data-stu-id="d69d7-103">Add comments in Excel</span></span>

<span data-ttu-id="d69d7-104">Cet exemple montre comment ajouter des commentaires à une cellule, y compris [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un collègue.</span><span class="sxs-lookup"><span data-stu-id="d69d7-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="d69d7-105">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="d69d7-105">Example scenario</span></span>

* <span data-ttu-id="d69d7-106">Le chef d’équipe maintient la planification du travail d’équipe.</span><span class="sxs-lookup"><span data-stu-id="d69d7-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="d69d7-107">Le chef d’équipe affecte un ID d’employé à l’enregistrement d’équipe.</span><span class="sxs-lookup"><span data-stu-id="d69d7-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="d69d7-108">Le chef d’équipe souhaite en informer l’employé.</span><span class="sxs-lookup"><span data-stu-id="d69d7-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="d69d7-109">En ajoutant un commentaire qui @mentions l’employé, un message personnalisé provenant de la feuille de calcul lui est envoyé par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="d69d7-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="d69d7-110">Par la suite, l’employé peut afficher le livre de travail et répondre au commentaire à sa convenance.</span><span class="sxs-lookup"><span data-stu-id="d69d7-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="d69d7-111">Solution</span><span class="sxs-lookup"><span data-stu-id="d69d7-111">Solution</span></span>

1. <span data-ttu-id="d69d7-112">Le script extrait les informations de l’employé de la feuille de calcul de l’employé.</span><span class="sxs-lookup"><span data-stu-id="d69d7-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="d69d7-113">Le script ajoute ensuite un commentaire (y compris l’e-mail de l’employé approprié) à la cellule appropriée dans l’enregistrement d’équipe.</span><span class="sxs-lookup"><span data-stu-id="d69d7-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="d69d7-114">Les commentaires existants dans la cellule sont supprimés avant d’ajouter le nouveau commentaire.</span><span class="sxs-lookup"><span data-stu-id="d69d7-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="d69d7-115">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="d69d7-115">Sample Excel file</span></span>

<span data-ttu-id="d69d7-116">Téléchargez <a href="excel-comments.xlsx">excel-comments.xlsx</a> pour un livre de travail prêt à l’emploi.</span><span class="sxs-lookup"><span data-stu-id="d69d7-116">Download <a href="excel-comments.xlsx">excel-comments.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="d69d7-117">Ajoutez le script suivant pour essayer l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="d69d7-117">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="d69d7-118">Exemple de code : ajouter des commentaires</span><span class="sxs-lookup"><span data-stu-id="d69d7-118">Sample code: Add comments</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);        
      
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```

## <a name="training-video-add-comments"></a><span data-ttu-id="d69d7-119">Vidéo de formation : ajouter des commentaires</span><span class="sxs-lookup"><span data-stu-id="d69d7-119">Training video: Add comments</span></span>

<span data-ttu-id="d69d7-120">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/CpR78nkaOFw).</span><span class="sxs-lookup"><span data-stu-id="d69d7-120">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
