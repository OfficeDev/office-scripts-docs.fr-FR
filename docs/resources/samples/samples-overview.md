---
title: Exemples de scripts Office
description: Exemples et scénarios De scripts Office disponibles.
ms.date: 08/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5798da37bd4166d18b41c005c4d8cc8a4b6c401d
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572485"
---
# <a name="office-scripts-samples-and-scenarios"></a>Exemples et scénarios de scripts Office

Cette section contient des solutions d’automatisation [basées sur des scripts Office](../../overview/excel.md) qui aident les utilisateurs finaux à réaliser l’automatisation des tâches quotidiennes. Il contient des scénarios réalistes auxquels les utilisateurs professionnels sont confrontés et fournit des solutions détaillées ainsi que des liens vidéo d’instructions pas à pas.

Pour chacun des projets de [base](#basics) et [au-delà des principes de base](#beyond-the-basics), consultez le code source, des [**vidéos YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0) pas à pas, et bien plus encore.

Dans [Scénarios](#scenarios), nous avons inclus quelques exemples de scénarios plus volumineux qui illustrent des cas d’usage réels.

Nous accueillons également les [contributions de la communauté](#community-contributions-and-fun-samples). Ces exemples sont open source.

> [!IMPORTANT]
> Veillez à respecter les prérequis pour les scripts Office avant d’essayer les exemples. La configuration requise pour votre abonnement et votre compte Microsoft 365 se trouve dans la [section « Configuration requise » de la vue d’ensemble des scripts Office pour Excel](../../overview/excel.md#requirements).

## <a name="basics"></a>Informations de base

| Projet | Détails |
|---------|---------|
| [Informations de base](excel-samples.md) | Ces exemples illustrent les blocs de construction fondamentaux des scripts Office. |
| [Ajouter des commentaires dans Excel](add-excel-comments.md) | Cet exemple ajoute des commentaires à une cellule, y compris @mentioning un collègue. |
| [Ajouter des images à un classeur](add-image-to-workbook.md) | Cet exemple ajoute une image à un classeur et copie une image sur plusieurs feuilles.|
| [Copier plusieurs tables Excel dans une table unique](copy-tables-combine.md) | Cet exemple combine les données de plusieurs tables Excel dans une table unique qui inclut toutes les lignes. |
| [Créer une table des matières de classeur](table-of-contents.md) | Cet exemple crée une table des matières avec des liens vers chaque feuille de calcul. |
| [Supprimer les filtres de la colonne du tableau](clear-table-filter-for-active-cell.md) | Cet exemple efface tous les filtres d’une colonne de table. |
| [Enregistrer les modifications quotidiennes dans Excel et les signaler à l’aide d’un flux Power Automate](report-day-to-day-changes.md) | Cet exemple utilise un flux Power Automate planifié pour enregistrer les lectures quotidiennes et signaler les modifications. |

## <a name="beyond-the-basics"></a>Notions intermédiaires

Consultez le projet de bout en bout suivant qui automatise des exemples de scénarios, ainsi que des scripts complets, des exemples de fichiers Excel utilisés et [des vidéos (hébergées sur YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Projet | Détails |
|---------|---------|
| [Combiner des feuilles de calcul dans un seul classeur](combine-worksheets-into-single-workbook.md) | Cet exemple utilise les scripts Office et Power Automate pour extraire des données d’autres classeurs dans un classeur unique. |
| [Convertir des fichiers CSV en classeurs Excel](convert-csv.md) | Cet exemple utilise les scripts Office et Power Automate pour créer des fichiers .xlsx à partir de fichiers .csv. |
| [Classeurs de référence croisée](excel-cross-reference.md) | Cet exemple utilise les scripts Office et Power Automate pour référencer et valider des informations dans différents classeurs. |
| [Compter les lignes vides dans une feuille spécifique ou dans toutes les feuilles](count-blank-rows.md) | Cet exemple détecte s’il existe des lignes vides dans les feuilles où vous prévoyez la présence de données, puis indique le nombre de lignes vides à utiliser dans un flux Power Automate. |
| [Email images de graphique et de tableau](email-images-chart-table.md) | Cet exemple utilise des scripts Office et des actions Power Automate pour créer un graphique et envoyer ce graphique en tant qu’image par e-mail. |
| [Appels d’extraction externes](external-fetch-calls.md) | Cet exemple permet d’obtenir des `fetch` informations à partir de GitHub pour le script. |
| [Gérer le mode de calcul dans Excel](excel-calculation.md) | Cet exemple montre comment utiliser le mode de calcul et calculer des méthodes dans Excel sur le Web à l’aide de scripts Office. |
| [Déplacer des lignes entre des tables](move-rows-across-tables.md) | Cet exemple montre comment déplacer des lignes entre des tables en enregistrant les filtres, puis en traitant et en réappliqueant les filtres. |
| [Générer des données Excel en tant que JSON](get-table-data.md) | Cette solution montre comment générer des données de table Excel en tant que JSON à utiliser dans Power Automate. |
| [Supprimer des liens hypertexte de chaque cellule d’une feuille de calcul Excel](remove-hyperlinks-from-cells.md) | Cet exemple efface tous les liens hypertexte de la feuille de calcul active. |
| [Exécuter un script sur tous les fichiers Excel d’un dossier](automate-tasks-on-all-excel-files-in-folder.md) | Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise (peut également être utilisé pour un dossier SharePoint). Il effectue des calculs sur les fichiers Excel, ajoute une mise en forme et insère un commentaire qui @mentions un collègue. |
| [Rédiger un grand ensemble de données](write-large-dataset.md) | Cet exemple montre comment envoyer une grande plage en tant que sous-plages plus petites. |

## <a name="scenarios"></a>Scénarios

Les scripts Office peuvent automatiser certaines parties de votre routine quotidienne. Ces tâches quotidiennes existent souvent dans des écosystèmes uniques, avec des classeurs Excel qui sont configurés de manière particulière. Ces exemples de scénarios plus volumineux illustrent de tels cas d’usage réels. Ils incluent à la fois les scripts Office et les classeurs, afin que vous puissiez voir le scénario de bout en bout.

| Scénario | Détails |
|---------|---------|
| [Analyser les téléchargements web](../scenarios/analyze-web-downloads.md) | Ce scénario comprend un script qui analyse les enregistrements de trafic web pour déterminer le pays d’origine d’un utilisateur. Il présente les compétences de l’analyse de texte, de l’utilisation de sous-fonctions dans les scripts, de l’application de la mise en forme conditionnelle et de l’utilisation de tables. |
| [Obtenir et représenter graphiquement les données du niveau d'eau auprès de la NOAA](../scenarios/noaa-data-fetch.md) | Ce scénario utilise un script Office pour extraire des données d’une source externe (base de données [NOAA Tides et Currents](https://tidesandcurrents.noaa.gov/)) et représenter les informations obtenues. Il met en évidence les compétences d’utilisation `fetch` pour obtenir des données et utiliser des graphiques. |
| [Calculatrice de notes](../scenarios/grade-calculator.md) | Ce scénario comporte un script qui valide l’enregistrement d’un instructeur pour les notes de leur classe. Il présente les compétences de la vérification des erreurs, de la mise en forme des cellules et des expressions régulières. |
| [Planifier des entretiens dans Teams](../scenarios/schedule-interviews-in-teams.md) | Ce scénario montre comment utiliser une feuille de calcul Excel pour gérer les heures de réunion des entretiens et établir un flux vers les planifications de réunions dans Teams. |
| [Rappels de tâches](../scenarios/task-reminders.md) | Ce scénario utilise un script Office dans un flux Power Automate pour envoyer des rappels à des collègues afin de mettre à jour l’état d’un projet. Il met en évidence les compétences de l’intégration de Power Automate et du transfert de données vers et depuis des scripts. |

## <a name="community-contributions-and-fun-samples"></a>Contributions de la communauté et exemples amusants

Nous accueillons les [contributions](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de notre communauté Office Scripts ! N’hésitez pas à créer une demande de tirage pour révision.

| Projet | Détails |
|---------|---------|
| [Jeu de la vie](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Le blog « Ready Player Zero » de Yutao Huang sur Excel Tech Community inclut un script pour modéliser [*Le jeu de la vie de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life) John Conway. |
| [Bouton Horloge perforée](../scenarios/punch-clock.md) | Ce script a été fourni par [Brian Gonzalez](https://github.com/b-gonzalez). Le scénario comprend un script et un bouton de script qui enregistre l’heure actuelle. |
| [Animation seasons greetings](community-seasons-greetings.md) | Ce script a été contribué par [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) dans l’esprit de la saison des fêtes! Il s’agit d’un script amusant qui montre un arbre de Noël chantant dans Excel sur le Web à l’aide de scripts Office. |

## <a name="leave-a-comment"></a>Laisser un commentaire

N’hésitez pas à laisser un commentaire, à faire une suggestion ou à enregistrer un problème à l’aide de la section **Commentaires** en bas de la page de documentation de l’exemple spécifique.
