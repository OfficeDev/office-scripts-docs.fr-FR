---
title: Office Exemples de scripts
description: Exemples Office scripts et scénarios disponibles.
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 625db792763606e8db77abdc4665b7db2732892f
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232738"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Exemples de scripts et scénarios

Cette section contient des Office [d’automatisation basées](../../overview/excel.md) sur des scripts qui aident les utilisateurs finaux à effectuer l’automatisation des tâches quotidiennes. Il contient des scénarios réalistes que les utilisateurs d’entreprise rencontrent et fournit des solutions détaillées, ainsi que des liens vidéo d’instructions pas à pas.

Pour chacun des [](#basics) projets de base et Au-delà des principes de [base,](#beyond-the-basics)consultez le code source, les vidéos [**YouTube**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)pas à pas et bien plus encore.

Dans [les scénarios,](#scenarios)nous avons inclus quelques exemples de scénarios plus importants qui montrent des cas d’utilisation réels.

Nous souhaitons également la [bienvenue aux contributions de la communauté.](#community-contributions)

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Informations de base

| Project | Détails |
|---------|---------|
| [Informations de base](../excel-samples.md) | Ces exemples montrent les blocs de construction fondamentaux pour Office scripts. |
| [Ajouter des commentaires dans Excel](add-excel-comments.md) | Cet exemple montre comment ajouter des commentaires à une cellule, y compris @mentioning collègue. |
| [Copier plusieurs Excel tables dans une seule table](copy-tables-combine.md) | Cet exemple combine les données de plusieurs tables Excel dans une seule table qui inclut toutes les lignes. |

## <a name="beyond-the-basics"></a>Notions intermédiaires

Consultez le projet de bout en bout suivant qui automatise des exemples de scénarios avec des scripts complets, des exemples de fichiers Excel utilisés et des vidéos (hébergées sur [YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Détails |
|---------|---------|
| [Compter les lignes vides dans une feuille spécifique ou dans toutes les feuilles](count-blank-rows.md) | Cet exemple détecte s’il existe des lignes vides dans les feuilles où vous prévoyez la présence de données, puis indique le nombre de lignes vides à utiliser dans un flux Power Automate données. |
| [Renvoi et mise en forme d’Excel fichier](excel-cross-reference.md) | Cette solution montre comment deux fichiers Excel peuvent être référencés et formatés à l’aide Office scripts et Power Automate. |
| [Images de tableau et de graphique de courrier électronique](email-images-chart-table.md) | Cet exemple utilise des Office scripts et des actions Power Automate pour créer un graphique et envoyer ce graphique en tant qu’image par courrier électronique. |
| [Appels de récupération externe](external-fetch-calls.md) | Cet exemple utilise `fetch` pour obtenir des informations à partir GitHub pour le script. |
| [Filtrer Excel tableau et obtenir une plage visible](filter-table-get-visible-range.md) | Cet exemple filtre un tableau Excel et renvoie la plage visible en tant qu’objet JSON. Ce JSON peut être fourni à un flux Power Automate dans le cadre d’une solution plus grande. |
| [Générer un identificateur unique dans un workbook](document-number-generator.md) | Ce scénario permet à un utilisateur de générer un numéro de document unique avec un format spécifique et d’ajouter une entrée à une plage ou un tableau. |
| [Gérer le mode de calcul dans Excel](excel-calculation.md) | Cet exemple montre comment utiliser le mode de calcul et calculer des méthodes dans Excel sur le Web à l’aide Office scripts. |
| [Déplacer des lignes dans des tableaux](move-rows-across-tables.md) | Cet exemple montre comment déplacer des lignes d’une table à l’autre en enregistrement des filtres, puis en traitant et réappliquent les filtres. |
| [Sortie Excel données en tant que JSON](get-table-data.md) | Cette solution indique comment Excel données de table en tant que JSON à utiliser dans Power Automate. |
| [Supprimer des liens hypertexte de chaque cellule d’une Excel de calcul](remove-hyperlinks-from-cells.md) | Cet exemple permet d’effacer tous les liens hypertexte de la feuille de calcul actuelle. |
| [Exécuter un script sur tous les fichiers Excel d’un dossier](automate-tasks-on-all-excel-files-in-folder.md) | Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise (peut également être utilisé pour un SharePoint dossier). Il effectue des calculs sur les fichiers Excel, ajoute une mise en forme et insère un commentaire qui @mentions un collègue. |
| [Envoyer une réunion Teams à partir de Excel données](send-teams-invite-from-excel-data.md) | Cette solution indique comment utiliser des scripts Office et des actions Power Automate pour sélectionner des lignes dans un fichier Excel et l’utiliser pour envoyer une invitation à une réunion Teams puis mettre à jour Excel. |

## <a name="scenarios"></a>Scénarios

Office Les scripts peuvent automatiser des parties de votre routine quotidienne. Ces tâches quotidiennes existent souvent dans des écosystèmes uniques, avec des Excel de travail qui sont spécifiquement mis en place. Ces exemples de scénarios plus importants montrent ces cas d’utilisation réels. Ils incluent les scripts Office et les workbooks, afin que vous pouvez voir le scénario de bout en bout.

| Scénario | Détails |
|---------|---------|
| [Analyser les téléchargements web](../scenarios/analyze-web-downloads.md) | Ce scénario comprend un script qui permet d’évaluer les enregistrements de trafic web pour déterminer le pays d’origine d’un utilisateur. Il présente les compétences de l’utilisation de sous-sections dans les scripts, de l’application de la mise en forme conditionnelle et de l’utilisation de tableaux. |
| [Obtenir et représenter graphiquement les données du niveau d'eau auprès de la NOAA](../scenarios/noaa-data-fetch.md) | Ce scénario utilise un script Office pour tirer des données à partir d’une source externe (base de données [NOAA - États-Unis](https://tidesandcurrents.noaa.gov/)et bases de données actuelles) et graphiquer les informations qui en résultent. Il met en évidence les compétences `fetch` d’utilisation pour obtenir des données et utiliser des graphiques. |
| [Calculatrice de notes](../scenarios/grade-calculator.md) | Ce scénario propose un script qui valide l’enregistrement d’un instructeur pour les notes de son cours. Il présente les compétences de vérification des erreurs, de mise en forme des cellules et d’expressions régulières. |
| [Rappels de tâche](../scenarios/task-reminders.md) | Ce scénario utilise un script Office dans un flux Power Automate pour envoyer des rappels à des collègues afin de mettre à jour l’état d’un projet. Il met en évidence les compétences de Power Automate’intégration et le transfert de données vers et depuis des scripts. |

## <a name="community-contributions"></a>Contributions de la communauté

Les [contributions](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) de notre communauté Office Scripts sont les bienvenues ! N’hésitez pas à créer une demande de tirer pour révision.

| Project | Détails |
|---------|---------|
| [Animation de message d’accueil de message d’accueil](community-seasons-greetings.md) | Ce script a été fourni par [Megan Black](https://www.linkedin.com/in/lesblackconsultant/) lors de la période des congés ! Il s’agit d’un script amusant qui montre une arborescence de Noël agréable dans Excel sur le Web à l’aide Office scripts. |

## <a name="try-it-out"></a>Try it out

Ces exemples sont open source. Essayez-les vous-même. Vous aurez besoin d’un compte scolaire ou scolaire Ou de travail Microsoft avec une licence d’Microsoft 365 (E3 ou supérieur). Il vous suffit de vous y rendre pour https://office.com vous inscrire à votre compte et commencer.

## <a name="leave-a-comment"></a>Laisser un commentaire

N’hésitez pas à laisser un commentaire, à faire une suggestion ou à enregistrer un problème à l’aide de la **section** Commentaires au bas de la page de documentation de l’exemple spécifique.
