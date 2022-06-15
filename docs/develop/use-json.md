---
title: Utiliser JSON pour transmettre des données vers et depuis Office Scripts
description: Découvrez comment structurer des données en objets JSON pour les utiliser avec des appels externes et des Power Automate
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 753097183a18f5d20ca2c78a3748c7a1d968ad42
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088151"
---
# <a name="use-json-to-pass-data-to-and-from-office-scripts"></a>Utiliser JSON pour transmettre des données vers et depuis Office Scripts

[JSON (JavaScript Object Notation)](https://www.w3schools.com/whatis/whatis_json.asp) est un format permettant de stocker et de transférer des données. Chaque objet JSON est une collection de paires nom/valeur qui peut être définie lors de la création. JSON est utile avec Office Scripts, car il peut gérer la complexité arbitraire des plages, des tables et d’autres modèles de données dans Excel. JSON vous permet d’analyser les données entrantes à partir de [services web](external-calls.md) et de transmettre des objets complexes via [des flux Power Automate](power-automate-integration.md).

Cet article se concentre sur l’utilisation de JSON avec Office Scripts. Nous vous recommandons d’abord d’en savoir plus sur le format à partir d’articles tels que [l’introduction JSON](https://www.w3schools.com/js/js_json_intro.asp) de W3 Schools.

## <a name="parse-json-data-into-a-range-or-table"></a>Analyser des données JSON dans une plage ou une table

Les tableaux d’objets JSON offrent un moyen cohérent de passer des lignes de données de table entre les applications et les services web. Dans ces cas, chaque objet JSON représente une ligne, tandis que les propriétés représentent les colonnes. Un script Office peut effectuer une boucle sur un tableau JSON et le réassemblage en tant que tableau 2D. Ce tableau est ensuite défini comme valeurs d’une plage et stocké dans un classeur. Les noms de propriété peuvent également être ajoutés en tant qu’en-têtes pour créer une table.

Le script suivant montre les données JSON converties en table. Notez que les données ne sont pas extraites d’une source externe. Cela est abordé plus loin dans cet article.

```typescript
/**
 * Sample JSON data. This would be replaced by external calls or
 * parameters getting data from Power Automate in a production script.
 */
const jsonData = [
  { "Action": "Edit", /* Action property with value of "Edit". */
    "N": 3370, /* N property with value of 3370. */
    "Percent": 17.85 /* Percent property with value of 17.85. */
  },
  // The rest of the object entries follow the same pattern.
  { "Action": "Paste", "N": 1171, "Percent": 6.2 },
  { "Action": "Clear", "N": 599, "Percent": 3.17 },
  { "Action": "Insert", "N": 352, "Percent": 1.86 },
  { "Action": "Delete", "N": 350, "Percent": 1.85 },
  { "Action": "Refresh", "N": 314, "Percent": 1.66 },
  { "Action": "Fill", "N": 286, "Percent": 1.51 },
];

/**
 * This script converts JSON data to an Excel table.
 */
function main(workbook: ExcelScript.Workbook) {
  // Create a new worksheet to store the imported data.
  const newSheet = workbook.addWorksheet();
  newSheet.activate();

  // Determine the data's shape by getting the properties in one object.
  // This assumes all the JSON objects have the same properties.
  const columnNames = getPropertiesFromJson(jsonData[0]);

  // Create the table headers using the property names.
  const headerRange = newSheet.getRangeByIndexes(0, 0, 1, columnNames.length);
  headerRange.setValues([columnNames]);

  // Create a new table with the headers.
  const newTable = newSheet.addTable(headerRange, true);

  // Add each object in the array of JSON objects to the table.
  const tableValues = jsonData.map(row => convertJsonToRow(row));
  newTable.addRows(-1, tableValues);
}

/**
 * This function turns a JSON object into an array to be used as a table row.
 */
function convertJsonToRow(obj: object) {
  const array: (string | number)[] = [];

  // Loop over each property and get the value. Their order will be the same as the column headers.
  for (let value in obj) {
    array.push(obj[value]);
  }
  return array;
}

/**
 * This function gets the property names from a single JSON object.
 */
function getPropertiesFromJson(obj: object) {
  const propertyArray: string[] = [];
  
  // Loop over each property in the object and store the property name in an array.
  for (let property in obj) {
    propertyArray.push(property);
  }

  return propertyArray;
}
```

> [!TIP]
> Si vous connaissez la structure du JSON, vous pouvez créer votre propre interface pour faciliter l’obtention de propriétés spécifiques. Vous pouvez remplacer les étapes de conversion JSON en tableau par des références de type sécurisé. L’extrait de code suivant montre ces étapes (maintenant commentées) remplacées par des appels qui utilisent une nouvelle `ActionRow` interface. Notez que cela rend la `convertJsonToRow` fonction inutile.
>
> ```typescript
>   // const tableValues = jsonData.map(row => convertJsonToRow(row));
>   // newTable.addRows(-1, tableValues);
>   // }
>
>      const actionRows: ActionRow[] = jsonData as ActionRow[];
>      // Add each object in the array of JSON objects to the table.
>      const tableValues = actionRows.map(row => [row.Action, row.N, row.Percent]);
>      newTable.addRows(-1, tableValues);
>    }
>    
>    interface ActionRow {
>      Action: string;
>      N: number;
>      Percent: number;
>    }
> ```

### <a name="get-json-data-from-external-sources"></a>Obtenir des données JSON à partir de sources externes

Il existe deux façons d’importer des données JSON dans votre classeur via un script Office.

- En tant que [paramètre](power-automate-integration.md#main-parameters-pass-data-to-a-script) avec un flux de Power Automate.
- Avec un `fetch` appel à un [service web externe](external-calls.md).

#### <a name="modify-the-sample-to-work-with-power-automate"></a>Modifier l’exemple pour qu’il fonctionne avec Power Automate

Les données JSON dans Power Automate peuvent être passées en tant que tableau d’objets génériques. Ajoutez une `object[]` propriété au script pour accepter ces données.

```typescript
// For Power Automate, replace the main signature in the previous sample with this one
// and remove the sample data.
function main(workbook: ExcelScript.Workbook, jsonData: object[]) {
```

Vous verrez ensuite une option dans le connecteur Power Automate à ajouter `jsonData` à l’action Exécuter le **script**.

:::image type="content" source="../images/json-parameter-power-automate.png" alt-text="Connecteur Excel Online (Entreprise) affichant une action exécuter un script avec le paramètre jsonData.":::

#### <a name="modify-the-sample-to-use-a-fetch-call"></a>Modifier l’exemple pour utiliser un `fetch` appel

Les services web peuvent répondre aux `fetch` appels avec des données JSON. Cela donne à votre script les données dont il a besoin tout en vous conservant dans Excel. En savoir plus sur `fetch` les appels externes et sur les appels [externes en lisant la prise en charge des appels d’API externes dans Office Scripts](external-calls.md).

```typescript
// For external services, replace the main signature in the previous sample with this one,
// add the fetch call, and remove the sample data.
async function main(workbook: ExcelScript.Workbook) {
  // Replace WEB_SERVICE_URL with the URL of whatever service you need to call.
  const response = await fetch('WEB_SERVICE_URL');
  const jsonData: object[] = await response.json();
```

## <a name="create-json-from-a-range"></a>Créer JSON à partir d’une plage

Les lignes et les colonnes d’une feuille de calcul impliquent souvent des relations entre leurs valeurs de données. Ligne d’une table mappée conceptuellement à un objet de programmation, chaque colonne étant une propriété de cet objet. Prenez en compte le tableau de données suivant. Chaque ligne représente une transaction enregistrée dans la feuille de calcul.

|ID |Date     |Montant |Fournisseur                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |43,54 $ |Best for you Organics Company |
|2  |6/3/2022 |67,23 $ |Liberty Bakery and Cafe       |
|3  |6/3/2022 |37,12 $ |Best for you Organics Company |
|4  |6/6/2022 |86,95 $ |Coho Vineyard                 |
|5  |6/7/2022 |13,64 $ |Liberty Bakery and Cafe       |

Chaque transaction (chaque ligne) a un ensemble de propriétés qui lui sont associées : « ID », « Date », « Amount » et « Vendor ». Cela peut être modélisé dans un script Office en tant qu’objet.

```typescript
// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

Les lignes de l’exemple de table correspondent aux propriétés de l’interface, de sorte qu’un script peut facilement convertir chaque ligne en objet `Transaction` . Cela est utile lors de la sortie des données pour Power Automate. Le script suivant itère sur chaque ligne de la table et l’ajoute à un `Transaction[]`.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="Sortie de la console du script précédent qui affiche les valeurs de propriété de l’objet.":::

### <a name="use-a-generic-object"></a>Utiliser un objet générique

L’exemple précédent suppose que les valeurs d’en-tête de table sont cohérentes. Si votre table contient des colonnes variables, vous devez créer un objet JSON générique. Le script suivant montre un script qui journalise n’importe quelle table en tant que JSON.

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Use the table header names as JSON properties.
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Get each data row in the table.
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  let jsonArray: object[] = [];

  // For each row, create a JSON object and assign each property to it based on the table headers.
  for (let i = 0; i < dataValues.length; i++) {
    // Create a blank generic JSON object.
    let jsonObject: { [key: string]: string } = {};
    for (let j = 0; j < dataValues[i].length; j++) {
      jsonObject[tableHeaders[j]] = dataValues[i][j] as string;
    }

    jsonArray.push(jsonObject);
  }

  // Do something with the objects, such as return them to a Power Automate flow.
  console.log(jsonArray);
}

```

## <a name="see-also"></a>Voir aussi

- [Prise en charge des appels d’API externes dans Scripts Office](external-calls.md)
- [Exemple : Utiliser des appels de récupération externe dans Office Scripts](../resources/samples/external-fetch-calls.md)
- [Exécuter des scripts Office avec Power Automate](power-automate-integration.md)