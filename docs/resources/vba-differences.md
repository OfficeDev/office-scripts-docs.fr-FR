---
title: Différences entre Office scripts et les macros VBA
description: Différences de comportement et d’API entre Office scripts et Excel macros VBA.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 53cd2d9b163a3d3c3f9ac9196b5f5126b539611a
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586016"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Différences entre Office scripts et les macros VBA

Office scripts et macros VBA ont beaucoup en commun. Elles permettent toutes deux aux utilisateurs d’automatiser des solutions via un enregistreur d’actions facile à utiliser et d’autoriser les modifications de ces enregistrements. Les deux frameworks sont conçus pour permettre aux personnes qui ne se considèrent pas comme des programmeurs de créer de petits programmes dans Excel.

La différence fondamentale est que les macros VBA sont développées pour les solutions de bureau et Office scripts sont conçus pour des solutions sécurisées basées sur le cloud. Actuellement, Office scripts sont uniquement pris en charge dans Excel sur le Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les domaines de mise au point pour Office solutions d’extensibilité différentes. Les scripts Office et les macros VBA sont conçus pour aider les utilisateurs finaux à créer des solutions, mais les scripts Office sont conçus pour le web et la collaboration (alors que VBA est destiné au bureau).":::

Cet article décrit les principales différences entre les macros VBA (ainsi que VBA en général) et Office scripts. Étant donné Office scripts sont uniquement disponibles pour Excel, il s’agit du seul hôte abordé ici.

## <a name="platform-and-ecosystem"></a>Plateforme et écosystème

VBA est pris en charge par Excel sur Windows mac. Office Scripts est pris en charge par Excel sur le Web.

Les deux solutions ont été conçues pour leurs plateformes respectives. VBA peut interagir avec le bureau d’un utilisateur pour se connecter à des technologies similaires, telles que COM et OLE. Toutefois, VBA ne dispose d’aucun moyen pratique pour appeler Internet. Office scripts utilisent un runtime universel pour JavaScript. Cela permet un comportement et une accessibilité cohérents, quel que soit l’ordinateur utilisé pour exécuter le script. Ils peuvent également effectuer des appels vers d’autres services web.

### <a name="script-support-for-excel-on-windows"></a>Prise en charge des scripts Excel sur Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>Sécurité

Les macros VBA ont la même habilitation de sécurité que Excel. Cela leur donne un accès complet à votre bureau. Office scripts ont uniquement accès au workbook, et non à l’ordinateur qui héberge le workbook. En outre, aucun jeton d’authentification JavaScript ne peut être partagé avec des scripts. Cela signifie que le script ne possède ni les jetons de l’utilisateur connexion, ni aucune fonctionnalité d’API pour la connexion à un service externe, de sorte qu’il ne peut pas utiliser les jetons existants pour effectuer des appels externes pour le compte de l’utilisateur.

Les administrateurs ont trois options pour les macros VBA : autoriser toutes les macros sur le client, n’autoriser aucune macro sur le client ou autoriser uniquement les macros avec des certificats signés. Ce manque de granularité rend difficile l’isolation d’un seul acteur mauvais. Actuellement, Office scripts peuvent être éteints pour un client entier, pour un client entier ou pour un groupe d’utilisateurs dans un client. Les administrateurs contrôlent également qui peut partager des scripts avec d’autres personnes et qui peut utiliser des scripts dans Power Automate.

## <a name="coverage"></a>Couverture

Actuellement, VBA offre une couverture plus complète des fonctionnalités Excel, en particulier celles disponibles sur le client de bureau. Office scripts couvrent presque tous les scénarios de Excel sur le Web. En outre, à mesure que de nouvelles fonctionnalités sont lancés sur le web, Office Scripts les prendra en charge pour l’enregistreur d’actions et les API JavaScript.

Office scripts ne peuvent pas Excel [événements](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) de niveau supérieur. Les scripts sont exécutés uniquement lorsqu’un utilisateur les démarre manuellement ou lorsqu’un flux Power Automate appelle le script.

## <a name="power-automate"></a>Power Automate

Office scripts peuvent être exécutés via Power Automate. Votre workbook peut être mis à jour par le biais de flux programmés ou pilotés par des événements, ce qui vous permet d’automatiser les flux de travail sans même ouvrir Excel. Cela signifie que tant que votre workbook est stocké dans OneDrive (et accessible à Power Automate), un flux peut exécuter vos scripts, que vous et votre organisation utilisez le client de bureau, Mac ou web de Excel.

VBA n’a pas de connecteur Power Automate de connexion. Tous les scénarios VBA pris en charge impliquent qu’un utilisateur participe à l’exécution de la macro.

Essayez les [scripts d’appel à partir d’un didacticiel Power Automate flux](../tutorials/excel-power-automate-manual.md) de travail pour commencer à apprendre à Power Automate. Vous pouvez également consulter l’exemple de [rappels de tâches automatisés](scenarios/task-reminders.md) pour voir Office scripts connectés à Teams via Power Automate dans un scénario réel.

## <a name="see-also"></a>Voir aussi

- [Office Scripts dans Excel sur le web](../overview/excel.md)
- [Exécuter Office scripts avec Power Automate](../develop/power-automate-integration.md)
- [Différences entre les scripts Office et les compléments Office](add-ins-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Référence VBA Excel](/office/vba/api/overview/excel)
