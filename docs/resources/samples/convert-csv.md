---
title: Convertir des fichiers CSV en Excel de travail
description: Découvrez comment utiliser Office scripts et Power Automate pour créer des fichiers .xlsx à partir .csv fichiers.
ms.date: 03/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 52619c1867b654fae3fce1a383a612f81f80d868
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585589"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Convertir des fichiers CSV en Excel de travail

De nombreux services exportent des données en tant que fichiers de valeurs séparées par des virgules (CSV). Cette solution automatise le processus de conversion de ces fichiers CSV en Excel de travail au format .xlsx format de fichier. Il utilise un flux [Power Automate](https://flow.microsoft.com) pour rechercher des fichiers avec l’extension .csv dans un dossier OneDrive et un script Office pour copier les données du fichier .csv dans un nouveau classeur Excel.

## <a name="solution"></a>Solution

1. Stockez les .csv et un fichier « Modèle » .xlsx vide dans un OneDrive dossier.
1. Créez un Office script pour analyser les données CSV dans une plage.
1. Créez un Power Automate pour lire les fichiers .csv et transmettre leur contenu au script.

## <a name="sample-files"></a>Exemples de fichiers

<a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true"> Téléchargezconvert-csv-example.zip</a> pour obtenir le fichier Template.xlsx et deux exemples .csv fichiers. Extrayez les fichiers dans un dossier de votre OneDrive. Cet exemple suppose que le dossier est nommé « output ».

Ajoutez le script suivant et créez un flux à l’aide des étapes données pour essayer l’exemple vous-même !

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Exemple de code : insérer des valeurs séparées par des virgules dans un workbook

```TypeScript
/**
 * Convert incoming CSV data into a range and add it to the workbook.
 */
function main(workbook: ExcelScript.Workbook, csv: string) {
  let sheet = workbook.getWorksheet("Sheet1");

  // Remove any Windows \r characters.
  csv = csv.replace(/\r/g, "");

  // Split each line into a row.
  let rows = csv.split("\n");
  /*
   * For each row, match the comma-separated sections.
   * For more information on how to use regular expressions to parse CSV files,
   * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
   */
  const csvMatchRegex = /(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g
  rows.forEach((value, index) => {
    if (value.length > 0) {
        let row = value.match(csvMatchRegex);
    
        // Check for blanks at the start of the row.
        if (row[0].charAt(0) === ',') {
          row.unshift("");
        }
    
        // Remove the preceding comma.
        row.forEach((cell, index) => {
          row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
        });
    
        // Create a 2D array with one row.
        let data: string[][] = [];
        data.push(row);
    
        // Put the data in the worksheet.
        let range = sheet.getRangeByIndexes(index, 0, 1, data[0].length);
        range.setValues(data);
    }
  });

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate flux : créer des fichiers .xlsx fichiers

1. [Connectez-Power Automate](https://flow.microsoft.com) et créez un **flux cloud programmé**.
1. Définissez le flux sur **Répéter tous les** « 1 » « Jour », puis sélectionnez **Créer**.
1. Obtenez le modèle Excel fichier. Il s’agit de la base de tous les fichiers .csv convertis. Ajoutez **une nouvelle étape** qui utilise **le connecteur OneDrive Entreprise** et l’action Obtenir **le contenu du** fichier. Indiquez le chemin d’accès au fichier « Template.xlsx ».
    * **Fichier** : /output/Template.xlsx
1. Renommez  l’étape Obtenir le contenu du fichier en allant dans le menu Pour obtenir le contenu **du fichier(...)** de cette étape (dans le coin supérieur droit du connecteur) et en sélectionnant l’option **Renommer**. Modifiez le nom de l’étape en « Obtenir Excel modèle ».

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="Le connecteur OneDrive Entreprise terminé dans Power Automate, renommé get Excel template.":::
1. Obtenez tous les fichiers dans le dossier « sortie ». Ajoutez **une nouvelle étape qui** utilise le **connecteur OneDrive Entreprise** et les fichiers de liste **dans l’action de** dossier. Fournissez le chemin d’accès du dossier qui contient .csv fichiers.
    * **Dossier** : /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate.":::
1. Ajoutez une condition de sorte que le flux fonctionne uniquement sur .csv fichiers. Ajoutez **une nouvelle étape** qui est le **contrôle condition** . Utilisez les valeurs suivantes pour la **condition**.
    * **Choisissez une valeur :** *Nom* (contenu dynamique des fichiers **de liste dans le dossier**). Notez que ce contenu dynamique a plusieurs résultats, donc   un contrôle Appliquer à chaque valeur entoure la **condition**.
    * **se termine par** (à partir de la liste liste liste)
    * **Choisissez une valeur** : .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="Contrôle Condition terminé avec l’application à chaque contrôle qui l’entoure.":::
1. Le reste du flux se trouve sous la section **Si** oui, car nous voulons uniquement agir sur .csv fichiers. Obtenez un fichier .csv individuel en ajoutant une nouvelle étape qui  utilise le connecteur **OneDrive Entreprise** et l’action Obtenir le contenu **du** fichier. Utilisez **l’ID du** contenu dynamique des fichiers **de liste dans le dossier**.
    * **Fichier** : *ID* (contenu dynamique des fichiers **de liste à l’étape du** dossier)
1. Renommez la nouvelle **étape Obtenir le contenu du** fichier en « Obtenir .csv fichier ». Cela permet de distinguer ce fichier du modèle Excel de données.
1. Faites du nouveau .xlsx, en utilisant le modèle Excel comme contenu de base. Ajoutez **une nouvelle étape** qui utilise **le connecteur OneDrive Entreprise** et l’action **Créer un** fichier. Utilisez les valeurs ci-après.
    * **Chemin d’accès du** dossier : /output
    * **Nom de** fichier *: nom sans extension*.xlsx (choisissez le nom sans contenu dynamique  *d’extension* dans les fichiers de liste du dossier et tapez manuellement « .xlsx » après celui-ci)
    * **Contenu du fichier** *: contenu de fichier* (contenu dynamique à partir **du modèle Obtenir Excel fichier**)

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Le fichier Obtenir .csv et créer des étapes de fichier du flux Power Automate’étapes.":::
1. Exécutez le script pour copier des données dans le nouveau workbook. Ajoutez **le connecteur Excel Online (Entreprise)** avec l’action **de script Exécuter**. Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier** : *ID* (contenu dynamique à partir **de créer un fichier**)
    * **Script** : convertir CSV
    * **csv** : *contenu de fichier* (contenu dynamique à partir de **Get .csv file**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="Le connecteur Excel Online (Entreprise) dans Power Automate.":::
1. Enregistrez le flux. Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre **onglet Mes flux** . N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.
1. Vous devez trouver de nouveaux .xlsx dans le dossier « sortie », ainsi que les fichiers .csv d’origine. Les nouveaux workbooks contiennent les mêmes données que les fichiers CSV.

## <a name="troubleshooting"></a>Résolution des problèmes

### <a name="script-testing"></a>Test de script

Pour tester le script sans utiliser Power Automate, affectez `csv` une valeur avant de l’utiliser. Essayez d’ajouter le code suivant en tant que première ligne de la `main` fonction et appuyez sur **Exécuter**.

```TypeScript
  csv = `1, 2, 3
         4, 5, 6
         7, 8, 9`;
```

### <a name="semicolon-separated-files-and-other-alternative-separators"></a>Fichiers séparés par des points-virgules et autres séparateurs

Certaines régions utilisent des points-virgules (';') pour séparer les valeurs des cellules au lieu de virgules. Dans ce cas, vous devez modifier les lignes suivantes dans le script.

1. Remplacez les virgules par des points-virgules dans l’instruction d’expression régulière. Cela commence par `let row = value.match`.

    ```TypeScript
    let row = value.match(/(?:;|\n|^)("(?:(?:"")*[^"]*)*"|[^";\n]*|(?:\n|$))/g);
    ```

1. Remplacez la virgule par un point-virgule dans la recherche de la première cellule vide. Cela commence par `if (row[0].charAt(0)`.

    ```TypeScript
    if (row[0].charAt(0) === ';') {
    ```

1. Remplacez la virgule par un point-virgule dans la ligne qui supprime le caractère de séparation du texte affiché. Cela commence par `row[index] = cell.indexOf`.

   ```TypeScript
      row[index] = cell.indexOf(";") === 0 ? cell.substr(1) : cell;
    ```

> [!NOTE]
> Si votre fichier utilise des tabulations ou tout autre caractère pour séparer les valeurs, `;` remplacez les substitutions `\t` ci-dessus par ou tout autre caractère utilisé.

### <a name="large-csv-files"></a>Fichiers CSV de grande taille

Si votre fichier possède des centaines de milliers de cellules, vous pouvez atteindre la [limite Excel transfert de données.](../../testing/platform-limits.md#excel) Vous devez forcer le script à se synchroniser avec Excel régulièrement. Le moyen le plus simple de le faire consiste à `console.log` appeler après le traitement d’un lot de lignes. Ajoutez les lignes de code suivantes pour y arriver.

1. Avant `rows.forEach((value, index) => {`, ajoutez la ligne suivante.

    ```TypeScript
      let rowCount = 0;
    ```

1. Après `range.setValues(data);`, ajoutez le code suivant. Notez que, en fonction du nombre de colonnes, `5000` vous devrez peut-être réduire ce nombre.

    ```TypeScript
      rowCount++;
      if (rowCount % 5000 === 0) {
        console.log("Syncing 5000 rows.");
      }
    ```

> [!WARNING]
> Si votre fichier CSV est très grand, vous pouvez avoir des problèmes de [délai](../../testing/platform-limits.md#power-automate) d’Power Automate. Vous devez diviser les données CSV en plusieurs fichiers avant de les convertir en Excel classeur.
