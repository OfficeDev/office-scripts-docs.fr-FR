---
title: Ajouter des images à un cahier de travail
description: Découvrez comment utiliser les scripts Office pour ajouter une image à un cahier de travail et la copier sur des feuilles.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 99c3cc2cacf6e535bdb882bb8414d23fd105be35
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52546028"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="407f4-103">Ajouter des images à un cahier de travail</span><span class="sxs-lookup"><span data-stu-id="407f4-103">Add images to a workbook</span></span>

<span data-ttu-id="407f4-104">Cet exemple montre comment travailler avec des images à l’aide d’un script Office en Excel.</span><span class="sxs-lookup"><span data-stu-id="407f4-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="407f4-105">Scénario</span><span class="sxs-lookup"><span data-stu-id="407f4-105">Scenario</span></span>

<span data-ttu-id="407f4-106">Les images aident à l’image de marque, à l’identité visuelle et aux modèles.</span><span class="sxs-lookup"><span data-stu-id="407f4-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="407f4-107">Ils aident à faire un cahier de travail plus qu’une table géante.</span><span class="sxs-lookup"><span data-stu-id="407f4-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="407f4-108">Le premier échantillon copie une image d’une feuille de travail à l’autre.</span><span class="sxs-lookup"><span data-stu-id="407f4-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="407f4-109">Cela pourrait être utilisé pour mettre le logo de votre entreprise dans la même position sur chaque feuille.</span><span class="sxs-lookup"><span data-stu-id="407f4-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="407f4-110">Le deuxième échantillon copie une image à partir d’une URL.</span><span class="sxs-lookup"><span data-stu-id="407f4-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="407f4-111">Cela pourrait être utilisé pour copier des photos qu’un collègue a stockées dans un dossier partagé à un cahier de travail connexe.</span><span class="sxs-lookup"><span data-stu-id="407f4-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="407f4-112">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="407f4-112">Sample Excel file</span></span>

<span data-ttu-id="407f4-113">Téléchargez le fichier <a href="add-images.xlsx">add-images.xlsxutilisé </a> dans ces échantillons et essayez-le vous-même!</span><span class="sxs-lookup"><span data-stu-id="407f4-113">Download the file <a href="add-images.xlsx">add-images.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="407f4-114">Exemple de code : Copiez une image sur des feuilles de travail</span><span class="sxs-lookup"><span data-stu-id="407f4-114">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="407f4-115">Exemple de code : Ajouter une image à partir d’une URL à un cahier de travail</span><span class="sxs-lookup"><span data-stu-id="407f4-115">Sample code: Add an image from a URL to a workbook</span></span>

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
