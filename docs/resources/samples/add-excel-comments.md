---
title: Ajouter des commentaires dans Excel
description: Découvrez comment utiliser Office scripts pour ajouter des commentaires dans une feuille de calcul.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: e5e5d17c076eceaf06fddeea1a67d31ee3581f31
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285933"
---
# <a name="add-comments-in-excel"></a><span data-ttu-id="4e954-103">Ajouter des commentaires dans Excel</span><span class="sxs-lookup"><span data-stu-id="4e954-103">Add comments in Excel</span></span>

<span data-ttu-id="4e954-104">Cet exemple montre comment ajouter des commentaires à une cellule, y compris [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un collègue.</span><span class="sxs-lookup"><span data-stu-id="4e954-104">This sample shows how to add comments to a cell including [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="4e954-105">Exemple de scénario</span><span class="sxs-lookup"><span data-stu-id="4e954-105">Example scenario</span></span>

* <span data-ttu-id="4e954-106">Le chef d’équipe maintient la planification du travail d’équipe.</span><span class="sxs-lookup"><span data-stu-id="4e954-106">The team lead maintains the shift schedule.</span></span> <span data-ttu-id="4e954-107">Le chef d’équipe affecte un ID d’employé à l’enregistrement d’équipe.</span><span class="sxs-lookup"><span data-stu-id="4e954-107">The team lead assigns an employee ID to the shift record.</span></span>
* <span data-ttu-id="4e954-108">Le chef d’équipe souhaite en informer l’employé.</span><span class="sxs-lookup"><span data-stu-id="4e954-108">The team lead wishes to notify the employee.</span></span> <span data-ttu-id="4e954-109">En ajoutant un commentaire qui @mentions l’employé, un message personnalisé provenant de la feuille de calcul lui est envoyé par courrier électronique.</span><span class="sxs-lookup"><span data-stu-id="4e954-109">By adding a comment that @mentions the employee, the employee is emailed with a custom message from the worksheet.</span></span>
* <span data-ttu-id="4e954-110">Par la suite, l’employé peut afficher le livre de travail et répondre au commentaire à sa convenance.</span><span class="sxs-lookup"><span data-stu-id="4e954-110">Subsequently, the employee can view the workbook and respond to the comment at their convenience.</span></span>

## <a name="solution"></a><span data-ttu-id="4e954-111">Solution</span><span class="sxs-lookup"><span data-stu-id="4e954-111">Solution</span></span>

1. <span data-ttu-id="4e954-112">Le script extrait les informations de l’employé de la feuille de calcul de l’employé.</span><span class="sxs-lookup"><span data-stu-id="4e954-112">The script extracts employee information from the employee worksheet.</span></span>
1. <span data-ttu-id="4e954-113">Le script ajoute ensuite un commentaire (y compris l’e-mail de l’employé approprié) à la cellule appropriée dans l’enregistrement d’équipe.</span><span class="sxs-lookup"><span data-stu-id="4e954-113">The script then adds a comment (including the relevant employee email) to the appropriate cell in the shift record.</span></span>
1. <span data-ttu-id="4e954-114">Les commentaires existants dans la cellule sont supprimés avant d’ajouter le nouveau commentaire.</span><span class="sxs-lookup"><span data-stu-id="4e954-114">Existing comments in the cell are removed before adding the new comment.</span></span>

## <a name="sample-code-add-comments"></a><span data-ttu-id="4e954-115">Exemple de code : ajouter des commentaires</span><span class="sxs-lookup"><span data-stu-id="4e954-115">Sample code: Add comments</span></span>

<span data-ttu-id="4e954-116">Téléchargez le fichier <a href="excel-comments.xlsx">excel-comments.xlsx</a> utilisé dans cet exemple et testez-le vous-même !</span><span class="sxs-lookup"><span data-stu-id="4e954-116">Download the file <a href="excel-comments.xlsx">excel-comments.xlsx</a> used in this sample and try it out yourself!</span></span>

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

## <a name="training-video-add-comments"></a><span data-ttu-id="4e954-117">Vidéo de formation : ajouter des commentaires</span><span class="sxs-lookup"><span data-stu-id="4e954-117">Training video: Add comments</span></span>

<span data-ttu-id="4e954-118">[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/CpR78nkaOFw).</span><span class="sxs-lookup"><span data-stu-id="4e954-118">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/CpR78nkaOFw).</span></span>
