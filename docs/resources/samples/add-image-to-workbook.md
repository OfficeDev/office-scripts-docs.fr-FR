---
title: Ajouter des images à un classeur
description: Découvrez comment utiliser Office scripts pour ajouter une image à un workbook et la copier sur plusieurs feuilles.
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 2ae77ca1295cf6e55443789506242d9cf888f9a1
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313903"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="43296-103">Ajouter des images à un classeur</span><span class="sxs-lookup"><span data-stu-id="43296-103">Add images to a workbook</span></span>

<span data-ttu-id="43296-104">Cet exemple montre comment utiliser des images à l’aide d’Office script dans Excel.</span><span class="sxs-lookup"><span data-stu-id="43296-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="43296-105">Scénario</span><span class="sxs-lookup"><span data-stu-id="43296-105">Scenario</span></span>

<span data-ttu-id="43296-106">Les images vous aident à utiliser la branding, l’identité visuelle et les modèles.</span><span class="sxs-lookup"><span data-stu-id="43296-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="43296-107">Ils contribuent à faire d’un workbook plus qu’un simple tableau.</span><span class="sxs-lookup"><span data-stu-id="43296-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="43296-108">Le premier exemple copie une image d’une feuille de calcul vers une autre.</span><span class="sxs-lookup"><span data-stu-id="43296-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="43296-109">Cela peut être utilisé pour placer le logo de votre entreprise à la même position sur chaque feuille.</span><span class="sxs-lookup"><span data-stu-id="43296-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="43296-110">Le deuxième exemple copie une image à partir d’une URL.</span><span class="sxs-lookup"><span data-stu-id="43296-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="43296-111">Cela peut être utilisé pour copier les photos qu’un collègue a stockées dans un dossier partagé dans un classeur associé.</span><span class="sxs-lookup"><span data-stu-id="43296-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="43296-112">Exemple Excel fichier</span><span class="sxs-lookup"><span data-stu-id="43296-112">Sample Excel file</span></span>

<span data-ttu-id="43296-113">Téléchargez <a href="add-images.xlsx">add-images.xlsx</a> pour un livre de travail prêt à l’emploi.</span><span class="sxs-lookup"><span data-stu-id="43296-113">Download <a href="add-images.xlsx">add-images.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="43296-114">Ajoutez les scripts suivants et essayez l’exemple vous-même !</span><span class="sxs-lookup"><span data-stu-id="43296-114">Add the following scripts and try the sample yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="43296-115">Exemple de code : copier une image sur plusieurs feuilles de calcul</span><span class="sxs-lookup"><span data-stu-id="43296-115">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="43296-116">Exemple de code : ajouter une image à partir d’une URL à un workbook</span><span class="sxs-lookup"><span data-stu-id="43296-116">Sample code: Add an image from a URL to a workbook</span></span>

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
