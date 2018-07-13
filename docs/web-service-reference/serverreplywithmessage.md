---
title: ServerReplyWithMessage
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- ServerReplyWithMessage
api_type:
- schema
ms.assetid: 113c6ff2-9592-44f0-b542-54e4d5122ccb
description: ServerReplyWithMessage 要素は、受信したメッセージへの応答として送信するのには、テンプレートのメッセージの ID を示します。
ms.openlocfilehash: f2d927ae18ac68523d4cdd173f0474fbbeb36c98
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19833390"
---
# <a name="serverreplywithmessage"></a><span data-ttu-id="d8c74-103">ServerReplyWithMessage</span><span class="sxs-lookup"><span data-stu-id="d8c74-103">ServerReplyWithMessage</span></span>

<span data-ttu-id="d8c74-104">**ServerReplyWithMessage**要素は、受信したメッセージへの応答として送信するのには、テンプレートのメッセージの ID を示します。</span><span class="sxs-lookup"><span data-stu-id="d8c74-104">The **ServerReplyWithMessage** element indicates the ID of the template message that is to be sent as a reply to incoming messages.</span></span> 
  
```XML
<ServerReplyWithMessage>
    <ItemId>
</ServerReplyWithMessage>
```

 <span data-ttu-id="d8c74-105">**ItemIdType**</span><span class="sxs-lookup"><span data-stu-id="d8c74-105">**ItemIdType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="d8c74-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="d8c74-106">Attributes and elements</span></span>

<span data-ttu-id="d8c74-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="d8c74-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="d8c74-108">属性</span><span class="sxs-lookup"><span data-stu-id="d8c74-108">Attributes</span></span>

<span data-ttu-id="d8c74-109">なし。</span><span class="sxs-lookup"><span data-stu-id="d8c74-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="d8c74-110">子要素</span><span class="sxs-lookup"><span data-stu-id="d8c74-110">Child elements</span></span>

|<span data-ttu-id="d8c74-111">**要素**</span><span class="sxs-lookup"><span data-stu-id="d8c74-111">**Element**</span></span>|<span data-ttu-id="d8c74-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="d8c74-112">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="d8c74-113">ItemId</span><span class="sxs-lookup"><span data-stu-id="d8c74-113">ItemId</span></span>](itemid.md) <br/> |<span data-ttu-id="d8c74-114">Exchange ストア内のアイテムの一意の識別子と変更キーを表します。</span><span class="sxs-lookup"><span data-stu-id="d8c74-114">Represents the unique identifier and change key of an item in the Exchange store.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="d8c74-115">親要素</span><span class="sxs-lookup"><span data-stu-id="d8c74-115">Parent elements</span></span>

|<span data-ttu-id="d8c74-116">**要素**</span><span class="sxs-lookup"><span data-stu-id="d8c74-116">**Element**</span></span>|<span data-ttu-id="d8c74-117">**説明**</span><span class="sxs-lookup"><span data-stu-id="d8c74-117">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="d8c74-118">アクション</span><span class="sxs-lookup"><span data-stu-id="d8c74-118">Actions</span></span>](actions.md) <br/> |<span data-ttu-id="d8c74-119">条件が満たされるときに、メッセージに対して実行される使用可能なアクションのセットを表します。</span><span class="sxs-lookup"><span data-stu-id="d8c74-119">Represents the set of actions that are available to be taken on a message when the conditions are fulfilled.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="d8c74-120">テキスト値</span><span class="sxs-lookup"><span data-stu-id="d8c74-120">Text value</span></span>

<span data-ttu-id="d8c74-121">なし。</span><span class="sxs-lookup"><span data-stu-id="d8c74-121">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="d8c74-122">備考</span><span class="sxs-lookup"><span data-stu-id="d8c74-122">Remarks</span></span>

<span data-ttu-id="d8c74-123">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="d8c74-123">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="d8c74-124">要素情報</span><span class="sxs-lookup"><span data-stu-id="d8c74-124">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="d8c74-125">名前空間</span><span class="sxs-lookup"><span data-stu-id="d8c74-125">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="d8c74-126">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="d8c74-126">Schema Name</span></span>  <br/> |<span data-ttu-id="d8c74-127">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="d8c74-127">Messages schema</span></span>  <br/> |
|<span data-ttu-id="d8c74-128">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="d8c74-128">Validation File</span></span>  <br/> |<span data-ttu-id="d8c74-129">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="d8c74-129">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="d8c74-130">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="d8c74-130">Can be Empty</span></span>  <br/> |<span data-ttu-id="d8c74-131">True</span><span class="sxs-lookup"><span data-stu-id="d8c74-131">True</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="d8c74-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="d8c74-132">See also</span></span>



- [<span data-ttu-id="d8c74-133">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="d8c74-133">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
