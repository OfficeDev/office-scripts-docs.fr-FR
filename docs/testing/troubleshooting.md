---
title: Dépannage de Office Scripts
description: Débogage des conseils et techniques pour les scripts Office, ainsi que des ressources d’aide.
ms.date: 10/30/2020
localization_priority: Normal
ms.openlocfilehash: b45957bd336edce527397253cacec8cb09df715a
ms.sourcegitcommit: 82d3c0ef1e187bcdeceb2b5fc3411186674fe150
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49342877"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="9776e-103">Dépannage de Office Scripts</span><span class="sxs-lookup"><span data-stu-id="9776e-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="9776e-104">Lorsque vous développez des scripts Office, vous pouvez faire des erreurs.</span><span class="sxs-lookup"><span data-stu-id="9776e-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="9776e-105">C'est bon.</span><span class="sxs-lookup"><span data-stu-id="9776e-105">It's okay.</span></span> <span data-ttu-id="9776e-106">Nous disposons d’outils qui permettent de trouver les problèmes et de faire fonctionner vos scripts parfaitement.</span><span class="sxs-lookup"><span data-stu-id="9776e-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="9776e-107">Journaux de console</span><span class="sxs-lookup"><span data-stu-id="9776e-107">Console logs</span></span>

<span data-ttu-id="9776e-108">Parfois, lors de la résolution des problèmes, vous voudrez imprimer des messages à l’écran.</span><span class="sxs-lookup"><span data-stu-id="9776e-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="9776e-109">Ces éléments peuvent vous indiquer la valeur actuelle des variables ou les chemins d’accès de code déclenchés.</span><span class="sxs-lookup"><span data-stu-id="9776e-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="9776e-110">Pour ce faire, consignez le texte dans la console.</span><span class="sxs-lookup"><span data-stu-id="9776e-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging myRange's address.");
console.log(myRange.getAddress());
```

<span data-ttu-id="9776e-111">Les chaînes transmises `console.log` s’afficheront dans la console de journalisation de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="9776e-111">Strings passed to `console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="9776e-112">Pour activer la console, appuyez sur le bouton de **sélection** et sélectionnez **logs...**</span><span class="sxs-lookup"><span data-stu-id="9776e-112">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="9776e-113">Les journaux n’ont pas d’incidence sur le classeur.</span><span class="sxs-lookup"><span data-stu-id="9776e-113">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="9776e-114">Messages d’erreur</span><span class="sxs-lookup"><span data-stu-id="9776e-114">Error messages</span></span>

<span data-ttu-id="9776e-115">Lorsque votre script Excel rencontre un problème, il génère une erreur.</span><span class="sxs-lookup"><span data-stu-id="9776e-115">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="9776e-116">Un message contextuel s’affiche pour vous demander si vous souhaitez **afficher les journaux**.</span><span class="sxs-lookup"><span data-stu-id="9776e-116">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="9776e-117">Appuyez sur ce bouton pour ouvrir la console et afficher les erreurs éventuelles.</span><span class="sxs-lookup"><span data-stu-id="9776e-117">Press that button to open the console and display any errors.</span></span>

## <a name="automate-tab-not-appearing-or-office-scripts-unavailable"></a><span data-ttu-id="9776e-118">L’onglet automatiser n’apparaît pas ou les scripts Office ne sont pas disponibles</span><span class="sxs-lookup"><span data-stu-id="9776e-118">Automate tab not appearing or Office Scripts unavailable</span></span>

<span data-ttu-id="9776e-119">Les étapes suivantes doivent vous aider à résoudre les problèmes liés à l’onglet **automatiser** qui n’apparaît pas dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="9776e-119">The following steps should help troubleshoot any problems related to the **Automate** tab not appearing in Excel on the web.</span></span>

1. <span data-ttu-id="9776e-120">Assurez-vous [que votre licence Microsoft 365 inclut des scripts Office](../overview/excel.md#requirements).</span><span class="sxs-lookup"><span data-stu-id="9776e-120">[Make sure your Microsoft 365 license includes Office Scripts](../overview/excel.md#requirements).</span></span>
1. <span data-ttu-id="9776e-121">[Vérifiez que votre navigateur est pris en charge](platform-limits.md#browser-support).</span><span class="sxs-lookup"><span data-stu-id="9776e-121">[Check that your browser is supported](platform-limits.md#browser-support).</span></span>
1. <span data-ttu-id="9776e-122">[Vérifiez que les cookies tiers sont activés](platform-limits.md#third-party-cookies).</span><span class="sxs-lookup"><span data-stu-id="9776e-122">[Ensure third-party cookies are enabled](platform-limits.md#third-party-cookies).</span></span>
1. <span data-ttu-id="9776e-123">[Assurez-vous que votre administrateur n’a pas désactivé les scripts Office dans le centre d’administration 365 de Microsoft](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="9776e-123">[Ensure that your admin has not disabled Office Scripts in the Microsoft 365 admin center](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## <a name="help-resources"></a><span data-ttu-id="9776e-124">Ressources d’aide</span><span class="sxs-lookup"><span data-stu-id="9776e-124">Help resources</span></span>

<span data-ttu-id="9776e-125">Le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts) est une communauté de développeurs souhaitant aider à coder les problèmes.</span><span class="sxs-lookup"><span data-stu-id="9776e-125">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="9776e-126">Souvent, vous pouvez trouver la solution à votre problème via une recherche de débordement de pile rapide.</span><span class="sxs-lookup"><span data-stu-id="9776e-126">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="9776e-127">Si ce n’est pas le cas, posez votre question et marquez-la à l’aide de la balise « Office-scripts ».</span><span class="sxs-lookup"><span data-stu-id="9776e-127">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="9776e-128">N’oubliez pas de mentionner que vous créez un *script* Office, et non un *complément* Office.</span><span class="sxs-lookup"><span data-stu-id="9776e-128">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="9776e-129">Si vous rencontrez un problème avec l’API JavaScript pour Office, créez un problème dans le référentiel GitHub [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="9776e-129">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="9776e-130">Les membres de l’équipe produit répondront aux problèmes et fourniront de l’aide.</span><span class="sxs-lookup"><span data-stu-id="9776e-130">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="9776e-131">La création d’un problème dans le référentiel **OfficeDev/Office-js** indique que vous avez trouvé un défaut dans la bibliothèque de l’API JavaScript Office que l’équipe produit doit résoudre.</span><span class="sxs-lookup"><span data-stu-id="9776e-131">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="9776e-132">En cas de problème avec l’enregistreur d’actions ou l’éditeur, envoyez des commentaires via le bouton **d’aide > commentaires** dans Excel.</span><span class="sxs-lookup"><span data-stu-id="9776e-132">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="9776e-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9776e-133">See also</span></span>

- [<span data-ttu-id="9776e-134">Office Scripts dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="9776e-134">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="9776e-135">Scripts de base pour les scripts Office dans Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="9776e-135">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="9776e-136">Limites des plateformes avec les scripts Office</span><span class="sxs-lookup"><span data-stu-id="9776e-136">Platform Limits with Office Scripts</span></span>](platform-limits.md)
- [<span data-ttu-id="9776e-137">Améliorer les performances de vos scripts Office</span><span class="sxs-lookup"><span data-stu-id="9776e-137">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="9776e-138">Annuler les effets d’un script Office</span><span class="sxs-lookup"><span data-stu-id="9776e-138">Undo the effects of an Office Script</span></span>](undo.md)
