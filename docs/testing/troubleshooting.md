---
title: Résolution des problèmes liés aux scripts Office
description: Débogage des conseils et techniques pour les scripts Office, ainsi que des ressources d’aide.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700204"
---
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="ab58d-103">Résolution des problèmes liés aux scripts Office</span><span class="sxs-lookup"><span data-stu-id="ab58d-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="ab58d-104">Lorsque vous développez des scripts Office, vous pouvez faire des erreurs.</span><span class="sxs-lookup"><span data-stu-id="ab58d-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="ab58d-105">C'est bon.</span><span class="sxs-lookup"><span data-stu-id="ab58d-105">It's okay.</span></span> <span data-ttu-id="ab58d-106">Nous disposons d’outils qui permettent de trouver les problèmes et de faire fonctionner vos scripts parfaitement.</span><span class="sxs-lookup"><span data-stu-id="ab58d-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="ab58d-107">Journaux de console</span><span class="sxs-lookup"><span data-stu-id="ab58d-107">Console logs</span></span>

<span data-ttu-id="ab58d-108">Parfois, lors de la résolution des problèmes, vous voudrez imprimer des messages à l’écran.</span><span class="sxs-lookup"><span data-stu-id="ab58d-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="ab58d-109">Ces éléments peuvent vous indiquer la valeur actuelle des variables ou les chemins d’accès de code déclenchés.</span><span class="sxs-lookup"><span data-stu-id="ab58d-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="ab58d-110">Pour ce faire, consignez le texte dans la console.</span><span class="sxs-lookup"><span data-stu-id="ab58d-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> <span data-ttu-id="ab58d-111">N’oubliez pas `load` d’utiliser les `sync` données de feuille de calcul et le classeur avant de consigner les propriétés de l’objet.</span><span class="sxs-lookup"><span data-stu-id="ab58d-111">Don't forget to `load` worksheet data and `sync` with the workbook before logging object properties.</span></span>

<span data-ttu-id="ab58d-112">Les chaînes transmises`console.log` s’afficheront dans la console de journalisation de l’éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="ab58d-112">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="ab58d-113">Pour activer la console, appuyez sur le bouton de **sélection** et sélectionnez **logs...**</span><span class="sxs-lookup"><span data-stu-id="ab58d-113">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="ab58d-114">Les journaux n’ont pas d’incidence sur le classeur.</span><span class="sxs-lookup"><span data-stu-id="ab58d-114">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="ab58d-115">Messages d’erreur</span><span class="sxs-lookup"><span data-stu-id="ab58d-115">Error messages</span></span>

<span data-ttu-id="ab58d-116">Lorsque votre script Excel rencontre un problème, il génère une erreur.</span><span class="sxs-lookup"><span data-stu-id="ab58d-116">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="ab58d-117">Un message contextuel s’affiche pour vous demander si vous souhaitez **afficher les journaux**.</span><span class="sxs-lookup"><span data-stu-id="ab58d-117">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="ab58d-118">Appuyez sur ce bouton pour ouvrir la console et afficher les erreurs éventuelles.</span><span class="sxs-lookup"><span data-stu-id="ab58d-118">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="ab58d-119">Ressources d’aide</span><span class="sxs-lookup"><span data-stu-id="ab58d-119">Help resources</span></span>

<span data-ttu-id="ab58d-120">Le [débordement de pile](https://stackoverflow.com/questions/tagged/office-scripts) est une communauté de développeurs souhaitant aider à coder les problèmes.</span><span class="sxs-lookup"><span data-stu-id="ab58d-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="ab58d-121">Souvent, vous pouvez trouver la solution à votre problème via une recherche de débordement de pile rapide.</span><span class="sxs-lookup"><span data-stu-id="ab58d-121">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="ab58d-122">Si ce n’est pas le cas, posez votre question et marquez-la à l’aide de la balise « Office-scripts ».</span><span class="sxs-lookup"><span data-stu-id="ab58d-122">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="ab58d-123">N’oubliez pas de mentionner que vous créez un *script*Office, et non un *complément*Office.</span><span class="sxs-lookup"><span data-stu-id="ab58d-123">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="ab58d-124">Si vous rencontrez un problème avec l’API JavaScript pour Office, créez un problème dans le référentiel GitHub [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="ab58d-124">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="ab58d-125">Les membres de l’équipe produit répondront aux problèmes et fourniront de l’aide.</span><span class="sxs-lookup"><span data-stu-id="ab58d-125">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="ab58d-126">La création d’un problème dans le référentiel **OfficeDev/Office-js** indique que vous avez trouvé un défaut dans la bibliothèque de l’API JavaScript Office que l’équipe produit doit résoudre.</span><span class="sxs-lookup"><span data-stu-id="ab58d-126">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="ab58d-127">En cas de problème avec l’enregistreur d’actions ou l’éditeur, envoyez des commentaires via le bouton **d’aide > commentaires** dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ab58d-127">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="ab58d-128">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ab58d-128">See also</span></span>

- [<span data-ttu-id="ab58d-129">Scripts Office dans Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="ab58d-129">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="ab58d-130">Scripts de base pour les scripts Office dans Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="ab58d-130">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="ab58d-131">Annuler les effets d’un script Office</span><span class="sxs-lookup"><span data-stu-id="ab58d-131">Undo the effects of an Office Script</span></span>](undo.md)
