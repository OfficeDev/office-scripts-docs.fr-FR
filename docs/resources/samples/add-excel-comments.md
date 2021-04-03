---
title: Ajouter des commentaires dans Excel
description: Découvrez comment utiliser les scripts Office pour ajouter des commentaires dans une feuille de calcul.
ms.date: 03/29/2021
localization_priority: Normal
ms.openlocfilehash: aaaf26df6973bd081290b0fbb67edecad8627e53
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571284"
---
# <a name="add-comments-in-excel"></a>Ajouter des commentaires dans Excel

Cet exemple montre comment ajouter des commentaires à une cellule, y compris [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) collègue.

## <a name="example-scenario"></a>Exemple de scénario

* Le chef d’équipe maintient la planification du travail d’équipe. Le chef d’équipe affecte un ID d’employé à l’enregistrement d’équipe.
* Le chef d’équipe souhaite en informer l’employé. En ajoutant un commentaire qui @mentions l’employé, un message personnalisé provenant de la feuille de calcul lui est envoyé par courrier électronique.
* Par la suite, l’employé peut afficher le livre de travail et répondre au commentaire à sa convenance.

## <a name="solution"></a>Solution

1. Le script extrait les informations de l’employé de la feuille de calcul de l’employé.
1. Le script ajoute ensuite un commentaire (y compris l’e-mail de l’employé approprié) à la cellule appropriée dans l’enregistrement d’équipe.
1. Les commentaires existants dans la cellule sont supprimés avant d’ajouter le nouveau commentaire.

## <a name="sample-code-add-comments"></a>Exemple de code : ajouter des commentaires

Téléchargez le fichier <a href="excel-comments.xlsx">excel-comments.xlsx</a> utilisé dans cet exemple et testez-le vous-même !

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
    console.log(employees); 

    const scheduleSheet = workbook.getWorksheet('Schedule');
    const table = scheduleSheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const scheduleData = range.getTexts();

    for (let i=0; i < scheduleData.length; i++) {
      let eId = scheduleData[i][3];

      let employeeInfo = employees.find(e => e[0] === eId);
      if (employeeInfo) {
        console.log("Found a match " + employeeInfo);
        let adminNotes = scheduleData[i][4];
        try { 
          let comment = workbook.getCommentByCell(range.getCell(i, 5));
          comment.delete();
        } catch {
            console.log("Ignore if there is no existing comment in the cell");
        }
        workbook.addComment(range.getCell(i,5), {
          mentions: [{
            email: employeeInfo[1],
            id: 0,
            name: employeeInfo[2]
          }],
          richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
        }, ExcelScript.ContentType.mention);        
        
      } else {
        console.log("No match for: " + eId);
      }
    }
    return;
}
```

## <a name="training-video-add-comments"></a>Vidéo de formation : ajouter des commentaires

[![Regardez une vidéo pas à pas sur l’ajout de commentaires dans un fichier Excel](../../images/comments-vid.jpg)](https://youtu.be/CpR78nkaOFw "Vidéo pas à pas sur l’ajout de commentaires dans un fichier Excel")
