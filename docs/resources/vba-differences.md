---
title: Différences entre Office scripts et les macros VBA
description: Différences de comportement et d’API entre Office scripts et Excel macros VBA.
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 612a5f21d935fd262a6e9fd12a3431956105636a
ms.sourcegitcommit: 4687693f02fc90a57ba30c461f35046e02e6f5fb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52545587"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Différences entre Office scripts et les macros VBA

Office Les scripts et les macros VBA ont beaucoup en commun. Elles permettent toutes deux aux utilisateurs d’automatiser des solutions via un enregistreur d’actions facile à utiliser et d’autoriser les modifications de ces enregistrements. Les deux frameworks sont conçus pour permettre aux personnes qui ne se considèrent pas comme des programmeurs de créer de petits programmes dans Excel.
La différence fondamentale est que les macros VBA sont développées pour les solutions de bureau et Office scripts sont conçus pour des solutions sécurisées basées sur le cloud. Actuellement, Office scripts sont uniquement pris en charge dans Excel sur le Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les domaines de mise au point pour Office solutions d’extensibilité différentes. Les scripts Office et les macros VBA sont conçus pour aider les utilisateurs finaux à créer des solutions, mais les scripts Office sont conçus pour le web et la collaboration (alors que VBA est destiné au bureau)":::

Cet article décrit les principales différences entre les macros VBA (ainsi que VBA en général) et Office scripts. Étant donné Office scripts sont uniquement disponibles pour Excel, il s’agit du seul hôte abordé ici.

## <a name="platform-and-ecosystem"></a>Plateforme et écosystème

VBA est conçu pour le bureau et Office scripts sont conçus pour le web. VBA peut interagir avec le bureau d’un utilisateur pour se connecter à des technologies similaires, telles que COM et OLE. Toutefois, VBA n’offre aucun moyen pratique de faire appel à Internet.

Office Les scripts utilisent un runtime universel pour JavaScript. Cela permet un comportement et une accessibilité cohérents, quel que soit l’ordinateur utilisé pour exécuter le script. Ils peuvent également effectuer des appels vers d’autres services web.

## <a name="security"></a>Sécurité

Les macros VBA ont la même habilitation de sécurité que Excel. Cela leur donne un accès complet à votre bureau. Office Les scripts ont uniquement accès au workbook, et non à l’ordinateur qui héberge le workbook. En outre, aucun jeton d’authentification JavaScript ne peut être partagé avec des scripts. Cela signifie que le script ne possède ni les jetons de l’utilisateur connexion, ni aucune fonctionnalité d’API pour la connexion à un service externe, de sorte qu’il ne peut pas utiliser les jetons existants pour effectuer des appels externes pour le compte de l’utilisateur.

Les administrateurs ont trois options pour les macros VBA : autoriser toutes les macros sur le client, n’autoriser aucune macro sur le client ou autoriser uniquement les macros avec des certificats signés. Ce manque de granularité rend difficile l’isolation d’un seul acteur mauvais. Actuellement, Office scripts sont soit en cours, soit éteints pour un client. Toutefois, nous travaillons pour donner aux administrateurs davantage de contrôle sur les scripts individuels et les créateurs de scripts.

## <a name="coverage"></a>Couverture

Actuellement, VBA offre une couverture plus complète des fonctionnalités Excel, en particulier celles disponibles sur le client de bureau. Office Les scripts couvrent presque tous les scénarios de Excel sur le Web. En outre, à mesure que de nouvelles fonctionnalités sont lancés sur le web, Office Scripts les prendra en charge pour l’enregistreur d’actions et les API JavaScript.

Office Les scripts ne sont pas Excel événements de [niveau supérieur.](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) Les scripts sont exécutés uniquement lorsqu’un utilisateur les démarre manuellement ou lorsqu’un flux Power Automate appelle le script.

## <a name="power-automate"></a>Power Automate

Office Les scripts peuvent être exécutés par le biais Power Automate. Votre workbook peut être mis à jour par le biais de flux programmés ou pilotés par des événements, ce qui vous permet d’automatiser les flux de travail sans même ouvrir Excel. Cela signifie que tant que votre workbook est stocké dans OneDrive (et accessible à Power Automate), un flux peut exécuter vos scripts, que vous et votre organisation utilisez le client de bureau, Mac ou web de Excel.

VBA n’a pas de connecteur Power Automate de connexion. Tous les scénarios VBA pris en charge impliquent qu’un utilisateur participe à l’exécution de la macro.

Essayez les scripts d’appel à partir d’un [didacticiel Power Automate flux](../tutorials/excel-power-automate-manual.md) de travail pour commencer à apprendre à Power Automate. Vous pouvez également consulter l’exemple de rappels de tâches [automatisés](scenarios/task-reminders.md) pour voir Office scripts connectés à Teams via Power Automate dans un scénario réel.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Exécuter Office scripts avec Power Automate](../develop/power-automate-integration.md)
- [Différences entre les scripts Office et les compléments Office](add-ins-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Référence VBA Excel](/office/vba/api/overview/excel)
