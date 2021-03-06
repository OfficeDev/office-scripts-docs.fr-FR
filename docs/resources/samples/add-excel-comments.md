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
# <a name="add-comments-in-excel"></a>Ajouter des commentaires dans Excel

Cet exemple montre comment ajouter des commentaires à une cellule, y compris [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) un collègue.

## <a name="example-scenario"></a>Exemple de scénario

* Le chef d’équipe maintient la planification du travail d’équipe. Le chef d’équipe affecte un ID d’employé à l’enregistrement d’équipe.
* Le chef d’équipe souhaite en informer l’employé. En ajoutant un commentaire qui @mentions l’employé, un message personnalisé provenant de la feuille de calcul lui est envoyé par courrier électronique.
* Par la suite, l’employé peut afficher le livre de travail et répondre au commentaire à sa convenance.

## <a name="solution"></a>Solution

1. Le script extrait les informations de l’employé de la feuille de calcul de l’employé.
1. Le script ajoute ensuite un commentaire (y compris l’e-mail de l’employé approprié) à la cellule appropriée dans l’enregistrement d’équipe.
1. Les commentaires existants dans la cellule sont supprimés avant d’ajouter le nouveau commentaire.

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez <a href="excel-comments.xlsx">excel-comments.xlsx</a> pour un livre de travail prêt à l’emploi. Ajoutez le script suivant pour essayer l’exemple vous-même !

## <a name="sample-code-add-comments"></a>Exemple de code : ajouter des commentaires

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

## <a name="training-video-add-comments"></a>Vidéo de formation : ajouter des commentaires

[Regardez Sudhi Genrethy parcourir cet exemple sur YouTube](https://youtu.be/CpR78nkaOFw).
