---
title: Différences entre les scripts Office et les macros VBA
description: Les différences de comportement et d’API entre les scripts Office et les macros VBA Excel.
ms.date: 06/30/2020
localization_priority: Normal
ms.openlocfilehash: 8c246545943341607a7aced4da792b8e49880cb0
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616688"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Différences entre les scripts Office et les macros VBA

Les scripts Office et les macros VBA ont beaucoup de choses en commun. Elles permettent aux utilisateurs d’automatiser des solutions via un enregistreur d’actions facile à utiliser et d’autoriser les modifications de ces enregistrements. Les deux infrastructures sont conçues pour aider les personnes qui ne peuvent pas prendre en compte les programmeurs à créer des petits programmes dans Excel.
La différence fondamentale réside dans le fait que les macros VBA sont développées pour les solutions de bureau et que les scripts Office sont conçus avec une prise en charge et une sécurité multiplateforme comme principes directeurs. Actuellement, les scripts Office sont uniquement pris en charge dans Excel sur le Web.

![Diagramme à quatre quadrants présentant les zones ciblées pour différentes solutions d’extensibilité Office. Les scripts Office et les macros VBA sont conçus pour aider les utilisateurs finaux à créer des solutions, mais les scripts Office sont créés pour le Web et la collaboration (tandis que VBA est destiné à l’ordinateur de bureau).)](../images/office-programmability-diagram.png)

Cet article décrit les principales différences entre les macros VBA (ainsi que VBA en général) et les scripts Office. Étant donné que les scripts Office sont disponibles uniquement pour Excel, c’est le seul hôte discuté ici.

## <a name="platform-and-ecosystem"></a>Plateforme et écosystème

VBA est conçu pour le bureau et les scripts Office sont conçus pour le Web. VBA peut interagir avec le Bureau d’un utilisateur pour se connecter à des technologies similaires, telles que COM et OLE. Toutefois, VBA n’a pas de moyen pratique d’appeler Internet.

Les scripts Office utilisent un Runtime universel pour JavaScript. Cela fournit un comportement cohérent et une accessibilité, indépendamment de l’ordinateur utilisé pour exécuter le script. Ils peuvent également passer des appels à d’autres services Web.

## <a name="security"></a>Sécurité

Les macros VBA ont le même habilitation de sécurité qu’Excel. Cela leur accorde un accès complet à votre bureau. Les scripts Office ont uniquement accès au classeur, et non à l’ordinateur qui héberge le classeur. De plus, il n’est pas possible de partager des jetons d’authentification JavaScript avec des scripts, de sorte que les scripts ne peuvent jamais s’authentifier auprès d’un service externe.

Les administrateurs disposent de trois options pour les macros VBA : autoriser toutes les macros sur le client, n’autoriser aucune macro sur le client ou n’autoriser que les macros avec des certificats signés. Ce manque de granularité rend difficile l’isolation d’un seul acteur incorrect. Actuellement, les scripts Office sont activés ou désactivés pour un client. Toutefois, nous travaillons pour permettre aux administrateurs de mieux contrôler les scripts et créateurs de script individuels.

## <a name="coverage"></a>Couverture

À l’heure actuelle, VBA offre une couverture plus complète des fonctionnalités Excel, en particulier celles disponibles sur le client de bureau. Les scripts Office couvrent presque tous les scénarios pour Excel sur le Web. De plus, en ce qui concerne les nouvelles fonctionnalités sur le Web, les scripts Office les prennent en charge à la fois pour l’enregistreur d’actions et les API JavaScript.

## <a name="power-automate"></a>Power Automate

Les scripts Office peuvent être exécutés via automate Power. Votre classeur peut être mis à jour par le biais de flux planifiés ou événementiels, ce qui vous permet d’automatiser les flux de travail sans même ouvrir Excel. Cela signifie que tant que votre classeur est stocké dans OneDrive (et accessible à Power Automated), un flux peut exécuter vos scripts, que vous-même et votre organisation utilisiez le client Excel de bureau, Mac ou Web.

VBA ne dispose pas d’un connecteur automate d’alimentation. Tous les scénarios VBA pris en charge impliquaient qu’un utilisateur participe à l’exécution de la macro.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Différences entre les scripts Office et les compléments Office](add-ins-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Référence VBA Excel](/office/vba/api/overview/excel)
