---
title: MarkAsJunk
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: f06bafc6-7ee3-4b2b-9fd1-7c51328f4729
description: MarkAsJunk 要素は、[迷惑メール] フォルダーにアイテムを移動して、送信者を受信拒否リストに追加する要求を指定します。
ms.openlocfilehash: fbb3eee7ce350954888931ca55b27f596656b161
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832350"
---
# <a name="markasjunk"></a><span data-ttu-id="e47d8-103">MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="e47d8-103">MarkAsJunk</span></span>

<span data-ttu-id="e47d8-104">**MarkAsJunk**要素は、[迷惑メール] フォルダーにアイテムを移動して、送信者を受信拒否リストに追加する要求を指定します。</span><span class="sxs-lookup"><span data-stu-id="e47d8-104">The **MarkAsJunk** element specifies the request to move an item to the junk mail folder and to add the sender to the blocked sender list.</span></span> 
  
```XML
<MarkAsJunk IsJunk="true | false" MoveItem="true | false">
   <ItemIds/>
</MarkAsJunk>
```

 <span data-ttu-id="e47d8-105">**MarkAsJunkType**</span><span class="sxs-lookup"><span data-stu-id="e47d8-105">**MarkAsJunkType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="e47d8-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="e47d8-106">Attributes and elements</span></span>

<span data-ttu-id="e47d8-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="e47d8-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="e47d8-108">属性</span><span class="sxs-lookup"><span data-stu-id="e47d8-108">Attributes</span></span>

|<span data-ttu-id="e47d8-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="e47d8-109">**Attribute**</span></span>|<span data-ttu-id="e47d8-110">**説明**</span><span class="sxs-lookup"><span data-stu-id="e47d8-110">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="e47d8-111">引数 IsJunk</span><span class="sxs-lookup"><span data-stu-id="e47d8-111">IsJunk</span></span>  <br/> |<span data-ttu-id="e47d8-112">**True** **引数 IsJunk**属性のテキスト値は、電子メールの送信者は、受信拒否リストに追加されていることを示します。</span><span class="sxs-lookup"><span data-stu-id="e47d8-112">A text value of **true** for the **IsJunk** attribute indicates that the email sender is added to the blocked sender list.</span></span> <span data-ttu-id="e47d8-113">**False**の値は、電子メール送信者の一覧に既に場合に、電子メールの送信者がブロックされる送信者の一覧から削除されているを示します。</span><span class="sxs-lookup"><span data-stu-id="e47d8-113">A value of **false** indicates that the email sender is removed from the blocked sender list, if the email sender is already on the list.</span></span>  <br/> |
|<span data-ttu-id="e47d8-114">MoveItem</span><span class="sxs-lookup"><span data-stu-id="e47d8-114">MoveItem</span></span>  <br/> |<span data-ttu-id="e47d8-115">の**場合は true** 、 **MoveItem**属性のテキスト値は、既定の迷惑メール フォルダーにアイテムを移動することを示します。</span><span class="sxs-lookup"><span data-stu-id="e47d8-115">A text value of **true** for the **MoveItem** attribute indicates that the item is moved to the default junk mail folder.</span></span> <span data-ttu-id="e47d8-116">**False**の値は、項目が既定の迷惑メール フォルダーに移動していないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e47d8-116">A value of **false** indicates that the item is not moved to the default junk mail folder.</span></span>  <br/> |
   
### <a name="child-elements"></a><span data-ttu-id="e47d8-117">子要素</span><span class="sxs-lookup"><span data-stu-id="e47d8-117">Child elements</span></span>

[<span data-ttu-id="e47d8-118">Itemid</span><span class="sxs-lookup"><span data-stu-id="e47d8-118">ItemIds</span></span>](itemids.md)
  
### <a name="parent-elements"></a><span data-ttu-id="e47d8-119">親要素</span><span class="sxs-lookup"><span data-stu-id="e47d8-119">Parent elements</span></span>

<span data-ttu-id="e47d8-120">なし。</span><span class="sxs-lookup"><span data-stu-id="e47d8-120">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="e47d8-121">備考</span><span class="sxs-lookup"><span data-stu-id="e47d8-121">Remarks</span></span>

<span data-ttu-id="e47d8-122">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="e47d8-122">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="e47d8-123">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="e47d8-123">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="e47d8-124">要素情報</span><span class="sxs-lookup"><span data-stu-id="e47d8-124">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="e47d8-125">名前空間</span><span class="sxs-lookup"><span data-stu-id="e47d8-125">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="e47d8-126">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="e47d8-126">Schema name</span></span>  <br/> |<span data-ttu-id="e47d8-127">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="e47d8-127">Messages schema</span></span>  <br/> |
|<span data-ttu-id="e47d8-128">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="e47d8-128">Validation file</span></span>  <br/> |<span data-ttu-id="e47d8-129">messages.xsd</span><span class="sxs-lookup"><span data-stu-id="e47d8-129">messages.xsd</span></span>  <br/> |
|<span data-ttu-id="e47d8-130">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="e47d8-130">Can be empty</span></span>  <br/> ||
   
