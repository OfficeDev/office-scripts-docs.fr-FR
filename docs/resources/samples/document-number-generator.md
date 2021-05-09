---
title: Générer un identificateur unique dans un workbook
description: Découvrez comment utiliser Office scripts pour générer un identificateur unique et ajouter une ligne à un tableau et une plage.
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 62c930bfc638dc46b36daf81b6d1ec976c90a8d0
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/08/2021
ms.locfileid: "52286084"
---
# <a name="generate-a-unique-identifier-in-a-workbook"></a>Générer un identificateur unique dans un workbook

Ce scénario permet à un utilisateur de générer un numéro de document unique avec un format spécifique et de l’ajouter en tant qu’entrée à une plage ou un tableau. La nouvelle entrée ou ligne ajoutée contiendra le numéro de document unique nouvellement généré et quelques autres attributs transmis au script.

Il existe deux versions de l’exemple pour ce scénario.

* [Version 1 : Lire et ajouter une ligne à une feuille de calcul contenant une plage simple](#sample-code-generate-key-and-add-row-to-range)

    _Avant l’ajout de la nouvelle ligne_

    :::image type="content" source="../../images/document-number-generator-range-before.png" alt-text="Feuille de calcul montrant une plage de données avant l’ajout d’une ligne":::

    _Une fois la nouvelle ligne ajoutée_

    :::image type="content" source="../../images/document-number-generator-range-after.png" alt-text="Feuille de calcul montrant une plage de données après l’ajout d’une ligne":::

* [Version 2 : lire et ajouter une ligne à un tableau](#sample-code-generate-key-and-add-row-to-table)

    _Avant l’ajout de la nouvelle ligne_

    :::image type="content" source="../../images/document-number-generator-table-before.png" alt-text="Feuille de calcul montrant un tableau avant l’ajout d’une ligne":::

    _Une fois la nouvelle ligne ajoutée_

    :::image type="content" source="../../images/document-number-generator-table-after.png" alt-text="Feuille de calcul montrant un tableau après l’ajout d’une ligne":::

## <a name="sample-excel-file"></a>Exemple Excel fichier

Téléchargez le fichier <a href="document-number-generator.xlsx">document-number-generator.xlsx</a> utilisé dans cette solution pour l’essayer vous-même !

## <a name="sample-code-generate-key-and-add-row-to-range"></a>Exemple de code : générer une clé et ajouter une ligne à une plage

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX  = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input:RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' && 
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('PlainSheet'); /* plain range sheet */
    const range = sheet.getUsedRange();

    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it. 
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length-1];

    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;
    
    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);
    
    // Get last row and compute next row address.
    const last = range.getLastRow();
    const target = last.getOffsetRange(1, 0);

    // Add a row with incoming data plus the computed key value.
    target.setValues([
      [
        nextKey, 
        /* Capitalize the document type. */
        input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
        input.documentName
      ]
    ])
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```

## <a name="sample-code-generate-key-and-add-row-to-table"></a>Exemple de code : générer une clé et ajouter une ligne au tableau

```TypeScript
function main(workbook: ExcelScript.Workbook, inputString: string): string {
    // Object to hold key prefixes for each document type.
    const PREFIX = {
        form: 'F',
        'work instruction': 'W'
    }

    // Length of the numeric part of the key.
    const KEYLENGTH = 6;

    // Parse the incoming string as object.
    const input: RequestData = JSON.parse(inputString);

    // Reject invalid request.
    if (input.docType.toLowerCase() !== 'form' &&
        input.docType.toLowerCase() !== 'work instruction') {
        throw `Invalid type sent to the script:  ${input.docType}. Should be one of the following: ${Object.keys(PREFIX)}`
    }

    // Get existing data in the sheet.
    const sheet = workbook.getWorksheet('TableSheet'); /* table sheet */
    const table = sheet.getTables()[0];
    const range = table.getRangeBetweenHeaderAndTotal();
    const data = range.getValues() as string[][];

    // Filter rows to match the incoming type and then extract the document number column (index 0) and then sort it.
    const selectIds = data.filter((value) => {
        return value[1].toLowerCase() === input.docType.toLowerCase();
    }).map((row) => row[0]).sort();

    // Get the max document ID for the type.
    const maxId = selectIds[selectIds.length - 1];


    // Extract numeric part.
    const numPart = maxId.substring(1);
    const nextNum = Number(numPart) + 1;

    // If we ever reach the max key value, throw an error.
    if (nextNum >= (10 ** KEYLENGTH)) {
        throw `Key sequence of ${nextNum} out of range for type: ${input.docType}.`
    }
    // Get the correct prefix value.
    const prefixVal: string = PREFIX[input.docType.toLowerCase()] as string;

    // Compute next key value.
    const nextKey = prefixVal + '0'.repeat(KEYLENGTH).substring(0, KEYLENGTH - String(nextNum).length) + String(nextNum);

    // Add a row with incoming data plus the computed key value.
    table.addRow(-1, [
            nextKey,
            /* Capitalize the document type. */
            input.docType[0].toUpperCase() + input.docType.toLowerCase().slice(1),
            input.documentName
        ]);
    console.log(`Added row: ${[nextKey, input.docType, input.documentName]}`)
    // Return the key value recorded in Excel.
    return nextKey;
}

// Incoming data structure.
interface RequestData {
    docType: string
    documentName: string
}
```
