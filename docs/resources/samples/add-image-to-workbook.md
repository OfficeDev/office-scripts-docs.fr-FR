---
title: Ajouter des images à un classeur
description: Découvrez comment utiliser Office scripts pour ajouter une image à un workbook et la copier sur plusieurs feuilles.
ms.date: 07/12/2021
localization_priority: Normal
ms.openlocfilehash: 993444aa328356f872db90d1b9d2403bf28be4de
ms.sourcegitcommit: a86b91c7e104bb7c26efd56de53b9e3976a34828
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/12/2021
ms.locfileid: "53394415"
---
# <a name="add-images-to-a-workbook"></a>Ajouter des images à un classeur

Cet exemple montre comment utiliser des images à l’aide d’Office script dans Excel.

## <a name="scenario"></a>Scénario

Les images vous aident à utiliser la branding, l’identité visuelle et les modèles. Ils aident à faire d’un workbook plus qu’une simple table de jeux.

Le premier exemple copie une image d’une feuille de calcul vers une autre. Cela peut être utilisé pour placer le logo de votre entreprise dans la même position sur chaque feuille.

Le deuxième exemple copie une image à partir d’une URL. Cela peut être utilisé pour copier les photos qu’un collègue a stockées dans un dossier partagé dans un classeur associé.

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez <a href="add-images.xlsx">add-images.xlsx</a> pour un livre de travail prêt à l’emploi. Ajoutez les scripts suivants et essayez l’exemple vous-même !

## <a name="sample-code-copy-an-image-across-worksheets"></a>Exemple de code : copier une image dans plusieurs feuilles de calcul

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Exemple de code : ajouter une image à partir d’une URL à un workbook

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
