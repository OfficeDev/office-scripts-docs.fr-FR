---
title: Convertir des fichiers CSV en Excel de travail
description: Découvrez comment utiliser des scripts Office et des Power Automate pour créer des .xlsx à partir .csv fichiers.
ms.date: 07/19/2021
localization_priority: Normal
ms.openlocfilehash: 0208453c1c60c5458e8f5e3978fe930d5b81377b90d402df738097c653665a2c
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2021
ms.locfileid: "57847502"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>Convertir des fichiers CSV en Excel de travail

De nombreux services exportent des données en tant que fichiers de valeurs séparées par des virgules (CSV). Cette solution automatise le processus de conversion de ces fichiers CSV en Excel au format .xlsx format de fichier. Il utilise un flux [Power Automate](https://flow.microsoft.com) pour rechercher des fichiers avec l’extension .csv dans un dossier OneDrive et un script Office pour copier les données du fichier .csv dans un nouveau classeur Excel.

## <a name="solution"></a>Solution

1. Stockez les .csv et un fichier « Modèle » .xlsx vide dans un OneDrive de données.
1. Créez un Office script pour analyser les données CSV dans une plage.
1. Créez un Power Automate pour lire les fichiers .csv et transmettre leur contenu au script.

## <a name="sample-files"></a>Exemples de fichiers

Téléchargez <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> pour obtenir le fichier Template.xlsx et deux exemples .csv fichiers. Extrayez les fichiers dans un dossier de votre OneDrive. Cet exemple suppose que le dossier est nommé « output ».

Ajoutez le script suivant et créez un flux à l’aide des étapes données pour essayer l’exemple vous-même !

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>Exemple de code : insérer des valeurs séparées par des virgules dans un workbook

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate flux : créer des fichiers .xlsx de données

1. Connectez-Power Automate et créez un **flux cloud programmé.** [](https://flow.microsoft.com)
1. Définissez le flux sur **Répéter tous les** « 1 » « Jour », puis sélectionnez **Créer.**
1. Obtenez le modèle Excel fichier. Il s’agit de la base de tous les fichiers .csv convertis. Ajoutez **une nouvelle étape** qui utilise le connecteur **OneDrive Entreprise** et l’action Obtenir le **contenu du** fichier. Indiquez le chemin d’accès au fichier « Template.xlsx ».
    * **Fichier**: /output/Template.xlsx
1. Renommez **l’étape** Obtenir le contenu du fichier en allant dans le menu **...** de cette étape (dans le coin supérieur droit du connecteur) et en sélectionnant l’option **Renommer.** Modifiez le nom de l’étape en « Obtenir Excel modèle ».

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="Le connecteur OneDrive Entreprise terminé dans Power Automate, renommé get Excel template.":::
1. Obtenez tous les fichiers dans le dossier « sortie ». Ajoutez **une nouvelle étape qui** utilise le connecteur **OneDrive Entreprise** et les fichiers de liste **dans l’action de** dossier. Fournissez le chemin d’accès du dossier qui contient .csv fichiers.
    * **Dossier**: /output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate.":::
1. Ajoutez une condition de sorte que le flux fonctionne uniquement sur .csv fichiers. Ajoutez **une nouvelle étape** qui est le contrôle **condition.** Utilisez les valeurs suivantes pour la **condition**.
    * **Choose a value**: *Name* (dynamic content from List files **in folder**). Notez que ce contenu dynamique   a plusieurs résultats, donc un contrôle Appliquer à chaque valeur entoure la **condition**.
    * **se termine par** (à partir de la liste liste
    * **Choisissez une valeur**: .csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="Contrôle condition terminé avec l’application à chaque contrôle qui l’entoure.":::
1. Le reste du flux se trouve sous la section **Si** oui, car nous voulons uniquement agir sur .csv fichiers. Obtenez un fichier .csv individuel  en ajoutant une nouvelle étape qui utilise le connecteur **OneDrive Entreprise** et l’action Obtenir le contenu **du** fichier. Utilisez **l’ID du** contenu dynamique des fichiers **de liste dans le dossier.**
    * **Fichier**: *ID* (contenu dynamique des fichiers **de liste à l’étape du** dossier)
1. Renommez la nouvelle **étape Obtenir le contenu du** fichier en « Obtenir .csv fichier ». Cela permet de distinguer ce fichier du modèle Excel de données.
1. Faites le nouveau fichier .xlsx, en utilisant le modèle Excel en tant que contenu de base. Ajoutez **une nouvelle étape** qui utilise le connecteur **OneDrive Entreprise** et l’action Créer **un** fichier. Utilisez les valeurs ci-après.
    * **Chemin d’accès du** dossier : /output
    * **Nom de fichier** *: nom sans extension*.xlsx (choisissez  le nom sans contenu dynamique *d’extension* dans les fichiers de liste du dossier et tapez manuellement « .xlsx » après celui-ci)
    * **Contenu du fichier**: *contenu de fichier* (contenu dynamique à partir du modèle Get **Excel)**

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="Le fichier Obtenir .csv et créer des étapes de fichier du flux Power Automate flux.":::
1. Exécutez le script pour copier des données dans le nouveau workbook. Ajoutez **le connecteur Excel Online (Entreprise)** avec l’action de script **Exécuter.** Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier**: *ID* (contenu dynamique à partir **de créer un fichier)**
    * **Script**: convertir CSV
    * **csv**: *contenu de fichier* (contenu dynamique à partir de Get .csv **fichier**)

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="Le connecteur Excel Online (Entreprise) dans Power Automate.":::
1. Enregistrez le flux. Utilisez le **bouton Test** dans la page d’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.
1. Vous devez trouver de nouveaux .xlsx dans le dossier « sortie », ainsi que les fichiers .csv d’origine. Les nouveaux workbooks contiennent les mêmes données que les fichiers CSV.
