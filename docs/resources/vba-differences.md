---
title: Différences entre les scripts Office et les macros VBA
description: Comportement et différences d’API entre les scripts Office et les macros VBA Excel.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 60e4fba6e63967302066f544b76fb20a8c8630a6
ms.sourcegitcommit: 34c7740c9bff0e4c7426e01029f967724bfee566
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/13/2022
ms.locfileid: "65393613"
---
# <a name="differences-between-office-scripts-and-vba-macros"></a>Différences entre les scripts Office et les macros VBA

Office les scripts et les macros VBA ont beaucoup en commun. Elles permettent aux utilisateurs d’automatiser des solutions via un enregistreur d’action facile à utiliser et d’autoriser les modifications de ces enregistrements. Les deux frameworks sont conçus pour permettre à des personnes qui ne se considèrent peut-être pas comme des programmeurs de créer de petits programmes dans Excel.

La différence fondamentale est que les macros VBA sont développées pour les solutions de bureau et Office scripts sont conçus pour des solutions sécurisées basées sur le cloud. Actuellement, les scripts Office sont pris en charge uniquement dans Excel sur le Web.

:::image type="content" source="../images/office-programmability-diagram.png" alt-text="Diagramme à quatre quadrants montrant les domaines d’intérêt des différentes solutions d’extensibilité Office. Les macros Office Scripts et VBA sont conçues pour aider les utilisateurs finaux à créer des solutions, mais Office scripts sont créés pour le web et la collaboration (alors que VBA est destiné au bureau).":::

Cet article décrit les principales différences entre les macros VBA (ainsi que VBA en général) et les scripts Office. Étant donné que Office scripts sont disponibles uniquement pour Excel, il s’agit du seul hôte abordé ici.

## <a name="platform-and-ecosystem"></a>Plateforme et écosystème

VBA est pris en charge par Excel sur Windows et Mac. Office Scripts est pris en charge par Excel sur le Web.

Les deux solutions ont été conçues pour leurs plateformes respectives. VBA peut interagir avec le bureau d’un utilisateur pour se connecter à des technologies similaires, telles que COM et OLE. Toutefois, VBA n’a pas de moyen pratique d’appeler à Internet. Office scripts utilisent un runtime universel pour JavaScript. Cela offre un comportement et une accessibilité cohérents, quel que soit l’ordinateur utilisé pour exécuter le script. Ils peuvent également passer des appels à d’autres services web.

### <a name="script-support-for-excel-on-windows"></a>Prise en charge des scripts pour Excel sur Windows

[!INCLUDE [Run-from-button support](../includes/run-from-button-desktop-support.md)]

## <a name="security"></a>Sécurité

Les macros VBA ont la même habilitation de sécurité que Excel. Cela leur donne un accès complet à votre bureau. Office scripts ont uniquement accès au classeur, et non à l’ordinateur hébergeant le classeur. En outre, aucun jeton d’authentification JavaScript ne peut être partagé avec des scripts. Cela signifie que le script n’a ni les jetons de l’utilisateur connecté, ni aucune fonctionnalité d’API pour la connexion à un service externe. Par conséquent, il ne peut pas utiliser les jetons existants pour effectuer des appels externes pour le compte de l’utilisateur.

Les administrateurs disposent de trois options pour les macros VBA : autoriser toutes les macros sur le locataire, n’autoriser aucune macro sur le locataire ou autoriser uniquement les macros avec des certificats signés. Ce manque de granularité rend difficile l’isolement d’un seul mauvais acteur. Actuellement, Office scripts peuvent être désactivés pour un locataire entier, activé pour un locataire entier ou activé pour un groupe d’utilisateurs dans un locataire. Les administrateurs contrôlent également qui peut partager des scripts avec d’autres personnes et qui peut utiliser des scripts dans Power Automate.

## <a name="coverage"></a>Couverture

Actuellement, VBA offre une couverture plus complète des fonctionnalités Excel, en particulier celles disponibles sur le client de bureau. Office scripts couvrent presque tous les scénarios de Excel sur le Web. En outre, comme de nouvelles fonctionnalités commencent sur le web, Office Scripts les prend en charge à la fois pour l’enregistreur d’actions et les API JavaScript.

Office scripts ne prennent pas en charge les [événements](/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-events-with-excel-objects) de niveau Excel. Les scripts sont exécutés uniquement lorsqu’un utilisateur les démarre manuellement ou lorsqu’un flux Power Automate appelle le script.

## <a name="power-automate"></a>Power Automate

Office scripts peuvent être exécutés via Power Automate. Votre classeur peut être mis à jour via des flux planifiés ou pilotés par des événements, ce qui vous permet d’automatiser les flux de travail sans même ouvrir Excel. Cela signifie que, tant que votre classeur est stocké dans OneDrive (et accessible à Power Automate), un flux peut exécuter vos scripts, que vous et votre organisation utilisiez le bureau, Mac ou le client web de Excel.

VBA n’a pas de connecteur Power Automate. Tous les scénarios VBA pris en charge impliquent un utilisateur participant à l’exécution de la macro.

Essayez les [scripts d’appel à partir d’un didacticiel de flux de Power Automate manuel](../tutorials/excel-power-automate-manual.md) pour commencer à en savoir plus sur Power Automate. Vous pouvez également consulter l’exemple de [rappels de tâches automatisés](scenarios/task-reminders.md) pour voir Office scripts connectés à Teams via Power Automate dans un scénario réel.

## <a name="see-also"></a>Voir aussi

- [scripts Office dans Excel](../overview/excel.md)
- [Exécuter des scripts Office avec Power Automate](../develop/power-automate-integration.md)
- [Différences entre les scripts Office et les compléments Office](add-ins-differences.md)
- [Dépannage de Office Scripts](../testing/troubleshooting.md)
- [Référence VBA Excel](/office/vba/api/overview/excel)
