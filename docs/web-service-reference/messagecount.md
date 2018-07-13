---
title: MessageCount
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MessageCount
api_type:
- schema
ms.assetid: 5efc5903-fcb3-44cf-aabb-b6912268df8e
description: MessageCount 要素には、現在のフォルダー内の会話の項目の合計数が含まれています。
ms.openlocfilehash: e598192d8a8eef84706de08245185c9b8d9a5b7f
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832453"
---
# <a name="messagecount"></a><span data-ttu-id="14a99-103">MessageCount</span><span class="sxs-lookup"><span data-stu-id="14a99-103">MessageCount</span></span>

<span data-ttu-id="14a99-104">**MessageCount**要素には、現在のフォルダー内の会話の項目の合計数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="14a99-104">The **MessageCount** element contains the total number of conversation items in the current folder.</span></span> 
  
[<span data-ttu-id="14a99-105">FindConversationResponse</span><span class="sxs-lookup"><span data-stu-id="14a99-105">FindConversationResponse</span></span>](findconversationresponse.md)
  
[<span data-ttu-id="14a99-106">スレッド</span><span class="sxs-lookup"><span data-stu-id="14a99-106">Conversations</span></span>](conversations-ex15websvcsotherref.md)
  
[<span data-ttu-id="14a99-107">会話 (ConversationType)</span><span class="sxs-lookup"><span data-stu-id="14a99-107">Conversation (ConversationType)</span></span>](conversation-conversationtype.md)
  
[<span data-ttu-id="14a99-108">MessageCount</span><span class="sxs-lookup"><span data-stu-id="14a99-108">MessageCount</span></span>](messagecount.md)
  
```XML
<MessageCount/>
```

 <span data-ttu-id="14a99-109">**xs:int**</span><span class="sxs-lookup"><span data-stu-id="14a99-109">**xs:int**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="14a99-110">属性および要素</span><span class="sxs-lookup"><span data-stu-id="14a99-110">Attributes and elements</span></span>

<span data-ttu-id="14a99-111">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="14a99-111">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="14a99-112">属性</span><span class="sxs-lookup"><span data-stu-id="14a99-112">Attributes</span></span>

<span data-ttu-id="14a99-113">なし。</span><span class="sxs-lookup"><span data-stu-id="14a99-113">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="14a99-114">子要素</span><span class="sxs-lookup"><span data-stu-id="14a99-114">Child elements</span></span>

<span data-ttu-id="14a99-115">なし。</span><span class="sxs-lookup"><span data-stu-id="14a99-115">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="14a99-116">親要素</span><span class="sxs-lookup"><span data-stu-id="14a99-116">Parent elements</span></span>

|<span data-ttu-id="14a99-117">**要素**</span><span class="sxs-lookup"><span data-stu-id="14a99-117">**Element**</span></span>|<span data-ttu-id="14a99-118">**説明**</span><span class="sxs-lookup"><span data-stu-id="14a99-118">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="14a99-119">会話 (ConversationType)</span><span class="sxs-lookup"><span data-stu-id="14a99-119">Conversation (ConversationType)</span></span>](conversation-conversationtype.md) <br/> |<span data-ttu-id="14a99-120">1 つのテーマを表します。</span><span class="sxs-lookup"><span data-stu-id="14a99-120">Represents a single conversation.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="14a99-121">テキスト値</span><span class="sxs-lookup"><span data-stu-id="14a99-121">Text value</span></span>

<span data-ttu-id="14a99-122">**MessageCount**要素のテキスト値は、テーマ フォルダー内のアイテムの合計数です。</span><span class="sxs-lookup"><span data-stu-id="14a99-122">The text value of the **MessageCount** element is the total number of conversation items in a folder.</span></span> 
  
## <a name="remarks"></a><span data-ttu-id="14a99-123">備考</span><span class="sxs-lookup"><span data-stu-id="14a99-123">Remarks</span></span>

<span data-ttu-id="14a99-124">この要素を記述するスキーマは、Exchange Web Services.This の要素がホストをで Exchange Server 2010 Service Pack 1 (SP1) を導入する IIS 仮想ディレクトリにあります。</span><span class="sxs-lookup"><span data-stu-id="14a99-124">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.This element was introduced in Exchange Server 2010 Service Pack 1 (SP1).</span></span>
  
## <a name="element-information"></a><span data-ttu-id="14a99-125">要素情報</span><span class="sxs-lookup"><span data-stu-id="14a99-125">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="14a99-126">名前空間</span><span class="sxs-lookup"><span data-stu-id="14a99-126">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="14a99-127">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="14a99-127">Schema name</span></span>  <br/> |<span data-ttu-id="14a99-128">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="14a99-128">Types schema</span></span>  <br/> |
|<span data-ttu-id="14a99-129">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="14a99-129">Validation file</span></span>  <br/> |<span data-ttu-id="14a99-130">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="14a99-130">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="14a99-131">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="14a99-131">Can be empty</span></span>  <br/> |<span data-ttu-id="14a99-132">False</span><span class="sxs-lookup"><span data-stu-id="14a99-132">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="14a99-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="14a99-133">See also</span></span>



<span data-ttu-id="14a99-134">
  [FindConversation 操作](findconversation-operation.md)</span><span class="sxs-lookup"><span data-stu-id="14a99-134">[FindConversation operation](findconversation-operation.md)</span></span>
  
[<span data-ttu-id="14a99-135">ApplyConversationAction 操作</span><span class="sxs-lookup"><span data-stu-id="14a99-135">ApplyConversationAction operation</span></span>](applyconversationaction-operation.md)


- [<span data-ttu-id="14a99-136">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="14a99-136">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)


[<span data-ttu-id="14a99-137">EWS での会話</span><span class="sxs-lookup"><span data-stu-id="14a99-137">Conversations in EWS</span></span>](http://msdn.microsoft.com/library/91e64629-db6c-4c94-9dcb-d386232e8467%28Office.15%29.aspx)
