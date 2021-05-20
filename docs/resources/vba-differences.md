---
title: Différences entre Office scripts et macros VBA
description: Les différences de comportement et d’API entre Office scripts et Excel macros VBA.
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 612a5f21d935fd262a6e9fd12a3431956105636a
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545587"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Différences entre Office scripts et macros VBA

Office Les scripts et les macros VBA ont beaucoup en commun. Ils permettent tous deux aux utilisateurs d’automatiser des solutions grâce à un enregistreur d’action facile à utiliser et permettent des modifications de ces enregistrements. Les deux cadres sont conçus pour permettre aux personnes qui ne se considèrent peut-être pas comme des programmeurs de créer de petits programmes Excel.
La différence fondamentale est que les macros VBA sont développées pour les solutions de bureau et Office scripts sont conçus pour des solutions sécurisées basées sur le cloud. Actuellement, les Office sont pris en charge uniquement dans Excel sur le Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Un diagramme à quatre quadrants montrant les domaines d’intérêt pour différentes solutions Office’extensibilité. Les Office scripts et les macros VBA sont conçus pour aider les utilisateurs finaux à créer des solutions, mais les scripts Office sont conçus pour le Web et la collaboration (alors que VBA est pour le bureau)":::

Cet article décrit les principales différences entre les macros VBA (ainsi que VBA en général) et les scripts Office. Puisque Office scripts ne sont disponibles que pour Excel, c’est le seul hôte en discussion ici.

## <a name="platform-and-ecosystem"></a>Plate-forme et écosystème

VBA est conçu pour le bureau et Office scripts sont conçus pour le Web. VBA peut interagir avec le bureau d’un utilisateur pour se connecter à des technologies similaires, telles que COM et OLE. Cependant, VBA n’a aucun moyen pratique d’appeler à l’Internet.

Office Les scripts utilisent un temps d’exécution universel pour JavaScript. Cela donne un comportement et une accessibilité cohérents, quelle que soit la machine utilisée pour exécuter le script. Ils peuvent également passer des appels vers d’autres services Web.

## <a name="security"></a>Sécurité

Les macros VBA ont la même habilitation de sécurité que Excel. Cela leur donne un accès complet à votre bureau. Office Les scripts n’ont accès qu’au cahier de travail, pas à la machine hébergeant le cahier de travail. En outre, aucun jeton d’authentification JavaScript ne peut être partagé avec des scripts. Cela signifie que le script n’a ni les jetons de l’utilisateur connecté ni aucune capacité d’API pour se connecter à un service externe, de sorte qu’ils ne sont pas en mesure d’utiliser les jetons existants pour faire des appels externes au nom de l’utilisateur.

Admins ont trois options pour les macros VBA: autoriser toutes les macros sur le locataire, ne pas autoriser de macros sur le locataire, ou ne permettre que des macros avec des certificats signés. Ce manque de granularité rend difficile l’isolement d’un seul mauvais acteur. Actuellement, Office scripts sont en cours ou en dehors pour un locataire. Cependant, nous travaillons à donner aux administrateurs plus de contrôle sur les scripts individuels et les créateurs de scripts.

## <a name="coverage"></a>couverture

Actuellement, VBA offre une couverture plus complète des fonctionnalités Excel, en particulier celles disponibles sur le client de bureau. Office Les scripts couvrent presque tous les scénarios pour Excel sur le Web. En outre, comme de nouvelles fonctionnalités débutent sur le web, Office scripts les prendra en charge à la fois pour l’enregistreur d’action et les API JavaScript.

Office Les scripts ne supportent pas Excel événements de niveau [supérieur](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects). Les scripts ne sont exécutés que lorsqu’un utilisateur les démarre manuellement ou lorsqu’Power Automate flux d’utilisateurs appelle le script.

## <a name="power-automate"></a>Power Automate

Office Les scripts peuvent être exécutés à travers Power Automate. Votre cahier de travail peut être mis à jour par des flux planifiés ou axés sur des événements, vous permettant d’automatiser les flux de travail sans même Excel. Cela signifie que tant que votre cahier de travail est stocké en OneDrive (et accessible à Power Automate), un flux peut exécuter vos scripts, que vous et votre organisation utilisiez le bureau, le Mac ou le client Web de Excel.

VBA n’a pas de connecteur Power Automate système. Tous les scénarios VBA pris en charge impliquent qu’un utilisateur s’occuper de l’exécution de la macro.

Essayez les [scripts d’appel à partir d’un Power Automate de flux](../tutorials/excel-power-automate-manual.md) manuel pour commencer à en apprendre davantage sur Power Automate. Vous pouvez également consulter [l’exemple des rappels de tâches automatisés](scenarios/task-reminders.md) pour voir les scripts Office connectés à Teams à Power Automate dans un scénario réel.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Exécutez Office scripts avec Power Automate](../develop/power-automate-integration.md)
- [Différences entre les scripts Office et les compléments Office](add-ins-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Référence VBA Excel](/office/vba/api/overview/excel)
