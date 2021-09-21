---
title: Combiner des workbooks dans un seul et même workbook
description: Découvrez comment utiliser Office scripts et Power Automate pour créer des feuilles de calcul de fusion à partir d’autres feuilles de calcul dans un seul et même workbook.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: ffb0fd13cf587184aec87ade36e5e0e661043b94
ms.sourcegitcommit: c23816babcc628b52f6d8aaa4b6342e04e83a5bd
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2021
ms.locfileid: "59460784"
---
# <a name="combine-worksheets-into-a-single-workbook"></a>Combiner des feuilles de calcul dans un seul classeur

Cet exemple montre comment tirer des données de plusieurs workbooks dans un seul et centralisé. Il utilise deux scripts : l’un pour récupérer des informations à partir d’un workbook et l’autre pour créer des feuilles de calcul avec ces informations. Il combine les scripts dans un flux Power Automate qui agit sur un dossier OneDrive entier.

> [!IMPORTANT]
> Cet exemple copie uniquement les valeurs des autresbooks. Il ne conserve pas la mise en forme, les graphiques, les tableaux ou d’autres objets.

## <a name="scenario"></a>Scénario

1. Créez un Excel dans votre OneDrive et ajoutez-y deux scripts à partir de cet exemple.
1. Créez un dossier dans votre OneDrive et ajoutez-y un ou plusieurs classeurs contenant des données.
1. Créez un flux pour obtenir tous les fichiers de ce dossier.
1. Utilisez le script **de données de feuille de** calcul Renvoyer pour obtenir les données à partir de chaque feuille de calcul dans chacun des workbooks.
1. Utilisez le script **Ajouter des feuilles de** calcul pour créer une feuille de calcul dans un seul et même workbook pour chaque feuille de calcul de tous les autres fichiers.

## <a name="sample-code-return-worksheet-data"></a>Exemple de code : renvoyer des données de feuille de calcul

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[]
{
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="sample-code-add-worksheets"></a>Exemple de code : ajouter des feuilles de calcul

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[])
{
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="power-automate-flow-combine-worksheets-into-a-single-workbook"></a>Power Automate flux : combiner des feuilles de calcul dans un seul et même workbook

1. Connectez-Power Automate et créez un flux **de cloud instantané.** [](https://flow.microsoft.com)
1. Sélectionnez **Déclencher manuellement un flux,** puis **sélectionnez Créer.**
1. Obtenez tous les fichiers du dossier. Dans cet exemple, nous allons utiliser un dossier nommé « output ». Ajoutez **une nouvelle étape qui** utilise le connecteur **OneDrive Entreprise** et les fichiers de liste **dans l’action de** dossier. Fournissez le chemin d’accès du dossier qui contient .csv fichiers.
    * **Dossier**: /output

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="Connecteur de OneDrive Entreprise terminé dans Power Automate.":::
1. Exécutez le script **de données de feuille de** calcul Return pour obtenir toutes les données de chacun des workbooks. Ajoutez **le connecteur Excel Online (Entreprise)** avec l’action de script **Exécuter.** Utilisez les valeurs suivantes pour l’action. Notez que lorsque vous ajoutez l’ID du fichier,  Power Automate encapsule l’action dans une application à chaque contrôle, afin que l’action soit effectuée sur chaque fichier. 
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier**: *ID* (contenu dynamique des fichiers **de liste dans le dossier)**
    * **Script**: renvoyer des données de feuille de calcul
1. Exécutez le script **Ajouter des feuilles de** calcul sur le nouveau Excel que vous avez créé. Cela permet d’ajouter les données de tous les autres workbooks. Après l’action de script  **Exécuter** précédente et à l’intérieur du contrôle Appliquer à chaque contrôle, ajoutez un connecteur Excel **Online (Entreprise)** avec l’action **Exécuter le script.** Utilisez les valeurs suivantes pour l’action.
    * **Emplacement** : OneDrive Entreprise
    * **Bibliothèque de documents** : OneDrive
    * **Fichier**: votre fichier
    * **Script**: ajouter des feuilles de calcul
    * **workbookName**: *Nom* (contenu dynamique des fichiers **de liste dans le dossier)**
    * **worksheetInformation** (après avoir  sélectionné le bouton Basculer vers l’ensemble du tableau, voir la remarque suivant l’image suivante) : résultat *(contenu* dynamique à partir du **script Exécuter)**

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="Les deux actions de script Exécuter à l’intérieur de l’application à chaque contrôle.":::
    > [!NOTE]
    > Sélectionnez **le bouton Basculer pour entrer l’intégralité** du tableau afin d’ajouter l’objet tableau directement, au lieu d’éléments individuels pour le tableau.
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="Bouton à basculer pour entrer un tableau entier dans une zone de saisie de champ de contrôle.":::
1. Enregistrez le flux. Utilisez le **bouton Test** sur la page de l’éditeur de flux ou exécutez le flux dans votre onglet **Mes flux.** N’oubliez pas d’autoriser l’accès lorsque vous y êtes invité.
1. Votre Excel doit maintenant avoir de nouvelles feuilles de calcul.
