---
title: Ajouter des images à un classeur
description: Découvrez comment utiliser les scripts Office pour ajouter une image à un classeur et la copier sur plusieurs feuilles.
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78c7779cf4d524ed62bf8d419135863228b23d33
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572604"
---
# <a name="add-images-to-a-workbook"></a>Ajouter des images à un classeur

Cet exemple montre comment utiliser des images à l’aide d’un script Office dans Excel.

## <a name="scenario"></a>Scénario

Les images facilitent la personnalisation, l’identité visuelle et les modèles. Ils aident à faire un classeur plus qu’une simple table géante.

Le premier exemple copie une image d’une feuille de calcul vers une autre. Cela peut être utilisé pour placer le logo de votre entreprise dans la même position sur chaque feuille.

Le deuxième exemple copie une image à partir d’une URL. Cela peut être utilisé pour copier des photos qu’un collègue a stockées dans un dossier partagé dans un classeur associé.

## <a name="sample-excel-file"></a>Exemple de fichier Excel

Téléchargez [add-images.xlsx](add-images.xlsx) pour un classeur prêt à l’emploi. Ajoutez les scripts suivants et essayez l’exemple vous-même !

## <a name="sample-code-copy-an-image-across-worksheets"></a>Exemple de code : Copier une image dans des feuilles de calcul

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>Exemple de code : Ajouter une image à partir d’une URL à un classeur

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
