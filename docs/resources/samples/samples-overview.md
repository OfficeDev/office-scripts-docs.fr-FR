---
title: Office Exemples de scripts
description: Disponible Office scripts et scénarios.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0ea9a8a8986681fca0e45784e2923c1d3b34576d
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545708"
---
# <a name="office-scripts-samples-and-scenarios"></a>Office Scripts échantillons et scénarios

Cette section contient des solutions [Office d’automatisation basées](../../overview/excel.md) sur scripts qui aident les utilisateurs finaux à automatiser les tâches quotidiennes. Il contient des scénarios réalistes que les utilisateurs professionnels font face et fournit des solutions détaillées ainsi que des liens vidéo pédagogiques étape par étape.

Pour chacun des projets dans [Basics](#basics) and [Beyond the basics](#beyond-the-basics), consultez le code source, les vidéos YouTube étape par [**étape,**](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)et plus encore.

Dans [Scénarios](#scenarios), nous avons inclus quelques échantillons de scénarios plus importants qui démontrent des cas d’utilisation réels.

Nous accueillons [également favorablement les contributions de la communauté](#community-contributions-and-fun-samples).

[!INCLUDE [Preview note](../../includes/preview-note.md)]

## <a name="basics"></a>Informations de base

| Project | Détails |
|---------|---------|
| [Informations de base](../excel-samples.md) | Ces échantillons démontrent des éléments constitutifs fondamentaux pour Office scripts. |
| [Ajouter des commentaires dans Excel](add-excel-comments.md) | Cet exemple ajoute des commentaires à une cellule, y compris @mentioning un collègue. |
| [Ajouter des images à un cahier de travail](add-image-to-workbook.md) | Cet exemple ajoute une image à un cahier de travail et copie une image à travers des feuilles.|
| [Copiez plusieurs tables Excel’entrée en une seule table](copy-tables-combine.md) | Cet exemple combine les données de plusieurs tables Excel en une seule table qui inclut toutes les lignes. |

## <a name="beyond-the-basics"></a>Notions intermédiaires

Découvrez le projet de bout en bout suivant qui automatise des exemples de scénarios ainsi que des scripts complets, des exemples de fichiers Excel utilisés et [des vidéos (hébergées sur YouTube).](https://www.youtube.com/playlist?list=PLr3zVPZrMOUMl88fs8uc2GGAePRnNe6m0)

| Project | Détails |
|---------|---------|
| [Comptez les lignes blanches dans une feuille spécifique ou dans toutes les feuilles](count-blank-rows.md) | Cet exemple détecte s’il y a des lignes blanches dans les feuilles où vous prévoyez que les données sont présentes, puis signalez le nombre de lignes blanches pour utilisation dans un flux Power Automate. |
| [Graphiques d’email et images de table](email-images-chart-table.md) | Cet exemple utilise Office scripts et Power Automate actions pour créer un graphique et envoyer ce graphique sous forme d’image par e-mail. |
| [Appels externes d’extraction](external-fetch-calls.md) | Cet exemple utilise `fetch` pour obtenir des informations de GitHub pour le script. |
| [Filtrez Excel table et obtenez la portée visible](filter-table-get-visible-range.md) | Cet exemple filtre une table Excel et renvoie la plage visible en tant qu’objet JSON. Ce JSON pourrait être fourni à un flux Power Automate dans le cadre d’une solution plus large. |
| [Gérer le mode de calcul en Excel](excel-calculation.md) | Cet exemple montre comment utiliser le mode de calcul et calculer les méthodes dans les Excel sur le Web’Office scripts. |
| [Déplacez les lignes à travers les tables](move-rows-across-tables.md) | Cet exemple montre comment déplacer les lignes à travers les tables en économisant des filtres, puis en traitant et en réappliquant les filtres. |
| [Sortie Excel données comme JSON](get-table-data.md) | Cette solution montre comment extrayant Excel de table comme JSON à utiliser dans Power Automate. |
| [Supprimer les hyperliens de chaque cellule dans une feuille Excel travail](remove-hyperlinks-from-cells.md) | Cet exemple efface tous les hyperliens de la feuille de travail actuelle. |
| [Exécuter un script sur tous les fichiers Excel d’un dossier](automate-tasks-on-all-excel-files-in-folder.md) | Ce projet effectue un ensemble de tâches d’automatisation sur tous les fichiers situés dans un dossier sur OneDrive Entreprise (peut également être utilisé pour un dossier SharePoint dossier). Il effectue des calculs sur les fichiers Excel, ajoute le formatage, et insère un commentaire qui @mentions un collègue. |
| [Ecrire un grand ensemble de données](write-large-dataset.md) | Cet exemple montre comment envoyer une grande plage sous forme de sous-ranges plus petites. |

## <a name="scenarios"></a>Scénarios

Office Les scripts peuvent automatiser certaines parties de votre routine quotidienne. Ces tâches quotidiennes existent souvent dans des écosystèmes uniques, avec des Excel de travail qui sont mis en place de manière particulière. Ces échantillons de scénarios plus importants démontrent de tels cas d’utilisation réels. Ils incluent à la fois Office scripts et les cahiers de travail, de sorte que vous pouvez voir le scénario de bout en bout.

| Scénario | Détails |
|---------|---------|
| [Analyser les téléchargements web](../scenarios/analyze-web-downloads.md) | Ce scénario comporte un script qui parse les enregistrements de trafic Web pour déterminer le pays d’origine d’un utilisateur. Il met en valeur les compétences de l’parsage de texte, en utilisant des sous-fonctions dans les scripts, l’application de formatage conditionnel, et de travailler avec des tables. |
| [Obtenir et représenter graphiquement les données du niveau d'eau auprès de la NOAA](../scenarios/noaa-data-fetch.md) | Ce scénario utilise un script Office pour extraire des données d’une source externe (la base [de données noaa Tides and Currents)](https://tidesandcurrents.noaa.gov/)et tracer les informations qui en résultent. Il met en évidence les compétences d’utilisation `fetch` pour obtenir des données et en utilisant des graphiques. |
| [Calculatrice de notes](../scenarios/grade-calculator.md) | Ce scénario comporte un script qui valide le dossier d’un instructeur pour les notes de leur classe. Il met en valeur les compétences de vérification des erreurs, de mise en forme cellulaire et d’expressions régulières. |
| [Rappels de tâches](../scenarios/task-reminders.md) | Ce scénario utilise un script Office dans un flux Power Automate pour envoyer des rappels à vos collègues pour mettre à jour l’état d’un projet. Il met en évidence les compétences de Power Automate’intégration et de transfert de données vers et depuis les scripts. |

## <a name="community-contributions-and-fun-samples"></a>Community contributions et des échantillons amusants

Nous [accueillons](https://github.com/OfficeDev/office-scripts-docs/blob/master/Contributing.md) favorablement les contributions de notre Office de scripts! N’hésitez pas à créer une demande d’examen pull.

| Project | Détails |
|---------|---------|
| [Jeu de la vie](https://techcommunity.microsoft.com/t5/excel-blog/ready-player-zero/ba-p/2246208) | Le blog « Ready Player Zero » de Yutao Huang sur le Excel Tech Community comprend un script pour modéliser [*Le Jeu de*](https://en.wikipedia.org/wiki/Conway%27s_Game_of_Life)la vie de John Conway . |
| [Animation de salutations saisons](community-seasons-greetings.md) | Ce scénario a été contribué par [Leslie Black dans](https://www.linkedin.com/in/lesblackconsultant/) l’esprit de la période des Fêtes! C’est un script amusant qui montre un arbre de Noël chantant dans Excel sur le Web utilisant Office Scripts. |

## <a name="try-it-out"></a>Try it out

Ces échantillons sont open source. Essayez-les vous-même. Vous aurez besoin d’un compte Microsoft de travail ou d’école du travail ou de l’école avec une licence pour Microsoft 365 abonnement (E3 ou plus). Il suffit de se https://office.com diriger vers pour se connecter à votre compte et commencer.

## <a name="leave-a-comment"></a>Laisser un commentaire

N’hésitez pas à laisser un commentaire, à faire une suggestion ou à enregistrer un problème en utilisant la section **Commentaires** au bas de la page de documentation de l’échantillon spécifique.
