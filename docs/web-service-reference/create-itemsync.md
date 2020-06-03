---
title: Create (ItemSync)
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- Create
api_type:
- schema
ms.assetid: cb5e64a2-66a5-4447-921e-7c13efb8f6bf
description: Create 要素は、ローカルのクライアントストアに作成する単一のアイテムを識別します。
ms.openlocfilehash: b9c0f28333594a6c17ee9581a227fc4773874fd6
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44460800"
---
# <a name="create-itemsync"></a><span data-ttu-id="3638c-103">Create (ItemSync)</span><span class="sxs-lookup"><span data-stu-id="3638c-103">Create (ItemSync)</span></span>

<span data-ttu-id="3638c-104">**Create**要素は、ローカルのクライアントストアに作成する単一のアイテムを識別します。</span><span class="sxs-lookup"><span data-stu-id="3638c-104">The **Create** element identifies a single item to create in the local client store.</span></span> 
  
- [<span data-ttu-id="3638c-105">SyncFolderItemsResponse</span><span class="sxs-lookup"><span data-stu-id="3638c-105">SyncFolderItemsResponse</span></span>](syncfolderitemsresponse.md) 
- [<span data-ttu-id="3638c-106">ResponseMessages</span><span class="sxs-lookup"><span data-stu-id="3638c-106">ResponseMessages</span></span>](responsemessages.md) 
- [<span data-ttu-id="3638c-107">SyncFolderItemsResponseMessage</span><span class="sxs-lookup"><span data-stu-id="3638c-107">SyncFolderItemsResponseMessage</span></span>](syncfolderitemsresponsemessage.md) 
- [<span data-ttu-id="3638c-108">変更 (アイテム)</span><span class="sxs-lookup"><span data-stu-id="3638c-108">Changes (Items)</span></span>](changes-items.md) 
- [<span data-ttu-id="3638c-109">Create (ItemSync)</span><span class="sxs-lookup"><span data-stu-id="3638c-109">Create (ItemSync)</span></span>](create-itemsync.md)
  
```xml
<Create>
   <Item/>
</Create>
```

```xml
<Create>
   <Task/> 
</Create>
```

```xml
<Create>
   <MeetingResponse/>
</Create>
```

```xml
<Create>
   <CalendarItem/>
</Create>
```

```xml
<Create>
   <MeetingMessage/>
</Create>
```

```xml
<Create>
   <DistributionList/>
</Create>
```

```xml
<Create>
   <MeetingCancellation/>
</Create>
```

```xml
<Create>
   <MeetingRequest/> 
</Create>
```

```xml
<Create>
   <Message/> 
</Create>
```

```xml
<Create>
   <Contact/> 
</Create>
```

<span data-ttu-id="3638c-110">**SyncFolderItemsCreateOrUpdateType**</span><span class="sxs-lookup"><span data-stu-id="3638c-110">**SyncFolderItemsCreateOrUpdateType**</span></span>

## <a name="attributes-and-elements"></a><span data-ttu-id="3638c-111">属性と要素</span><span class="sxs-lookup"><span data-stu-id="3638c-111">Attributes and elements</span></span>

<span data-ttu-id="3638c-112">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="3638c-112">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="3638c-113">属性</span><span class="sxs-lookup"><span data-stu-id="3638c-113">Attributes</span></span>

<span data-ttu-id="3638c-114">なし。</span><span class="sxs-lookup"><span data-stu-id="3638c-114">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="3638c-115">子要素</span><span class="sxs-lookup"><span data-stu-id="3638c-115">Child elements</span></span>

|<span data-ttu-id="3638c-116">**Element**</span><span class="sxs-lookup"><span data-stu-id="3638c-116">**Element**</span></span>|<span data-ttu-id="3638c-117">**説明**</span><span class="sxs-lookup"><span data-stu-id="3638c-117">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="3638c-118">Item</span><span class="sxs-lookup"><span data-stu-id="3638c-118">Item</span></span>](item.md) <br/> |<span data-ttu-id="3638c-119">作成する汎用の Exchange アイテムを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-119">Represents a generic Exchange item to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-120">Message</span><span class="sxs-lookup"><span data-stu-id="3638c-120">Message</span></span>](message-ex15websvcsotherref.md) <br/> |<span data-ttu-id="3638c-121">作成する Exchange 電子メールメッセージを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-121">Represents an Exchange e-mail message to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-122">CalendarItem</span><span class="sxs-lookup"><span data-stu-id="3638c-122">CalendarItem</span></span>](calendaritem.md) <br/> |<span data-ttu-id="3638c-123">作成する Exchange 予定表アイテムを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-123">Represents an Exchange calendar item to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-124">Contact</span><span class="sxs-lookup"><span data-stu-id="3638c-124">Contact</span></span>](contact.md) <br/> |<span data-ttu-id="3638c-125">作成する Exchange の連絡先アイテムを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-125">Represents an Exchange contact item to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-126">DistributionList</span><span class="sxs-lookup"><span data-stu-id="3638c-126">DistributionList</span></span>](distributionlist.md) <br/> |<span data-ttu-id="3638c-127">作成する配布リストを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-127">Represents a distribution list to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-128">MeetingMessage</span><span class="sxs-lookup"><span data-stu-id="3638c-128">MeetingMessage</span></span>](meetingmessage.md) <br/> |<span data-ttu-id="3638c-129">作成する会議メッセージを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-129">Represents a meeting message to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-130">MeetingRequest</span><span class="sxs-lookup"><span data-stu-id="3638c-130">MeetingRequest</span></span>](meetingrequest.md) <br/> |<span data-ttu-id="3638c-131">作成する会議出席依頼を表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-131">Represents a meeting request to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-132">MeetingResponse</span><span class="sxs-lookup"><span data-stu-id="3638c-132">MeetingResponse</span></span>](meetingresponse.md) <br/> |<span data-ttu-id="3638c-133">作成する会議の応答を表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-133">Represents a meeting response to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-134">MeetingCancellation</span><span class="sxs-lookup"><span data-stu-id="3638c-134">MeetingCancellation</span></span>](meetingcancellation.md) <br/> |<span data-ttu-id="3638c-135">作成する会議の取り消しを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-135">Represents a meeting cancellation to create.</span></span>  <br/> |
|[<span data-ttu-id="3638c-136">タスク</span><span class="sxs-lookup"><span data-stu-id="3638c-136">Task</span></span>](task.md) <br/> |<span data-ttu-id="3638c-137">作成するタスクを表します。</span><span class="sxs-lookup"><span data-stu-id="3638c-137">Represents a task to create.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="3638c-138">親要素</span><span class="sxs-lookup"><span data-stu-id="3638c-138">Parent elements</span></span>

|<span data-ttu-id="3638c-139">**要素**</span><span class="sxs-lookup"><span data-stu-id="3638c-139">**Element**</span></span>|<span data-ttu-id="3638c-140">**説明**</span><span class="sxs-lookup"><span data-stu-id="3638c-140">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="3638c-141">変更 (アイテム)</span><span class="sxs-lookup"><span data-stu-id="3638c-141">Changes (Items)</span></span>](changes-items.md) <br/> |<span data-ttu-id="3638c-142">クライアント上のアイテムと Exchange サーバー上のアイテムの間の相違点の種類を表す、変更の種類のシーケンス配列を含みます。</span><span class="sxs-lookup"><span data-stu-id="3638c-142">Contains a sequence array of change types that represent the types of differences between the items on the client and the items on the Exchange server.</span></span>  <br/> |
   
## <a name="remarks"></a><span data-ttu-id="3638c-143">注釈</span><span class="sxs-lookup"><span data-stu-id="3638c-143">Remarks</span></span>

<span data-ttu-id="3638c-144">この要素を記述するスキーマは、クライアントアクセスサーバーの役割がインストールされている Microsoft Exchange Server 2007 を実行しているコンピューターの EWS 仮想ディレクトリにあります。</span><span class="sxs-lookup"><span data-stu-id="3638c-144">The schema that describes this element is located in the EWS virtual directory of the computer that is running MicrosoftExchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="3638c-145">要素の情報</span><span class="sxs-lookup"><span data-stu-id="3638c-145">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="3638c-146">Namespace</span><span class="sxs-lookup"><span data-stu-id="3638c-146">Namespace</span></span>  <br/> |https://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="3638c-147">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="3638c-147">Schema name</span></span>  <br/> |<span data-ttu-id="3638c-148">Types スキーマ</span><span class="sxs-lookup"><span data-stu-id="3638c-148">Types schema</span></span>  <br/> |
|<span data-ttu-id="3638c-149">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="3638c-149">Validation file</span></span>  <br/> |<span data-ttu-id="3638c-150">型 .xsd</span><span class="sxs-lookup"><span data-stu-id="3638c-150">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="3638c-151">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="3638c-151">Can be empty</span></span>  <br/> |<span data-ttu-id="3638c-152">正しくない</span><span class="sxs-lookup"><span data-stu-id="3638c-152">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="3638c-153">関連項目</span><span class="sxs-lookup"><span data-stu-id="3638c-153">See also</span></span>

- [<span data-ttu-id="3638c-154">SyncFolderItems 操作</span><span class="sxs-lookup"><span data-stu-id="3638c-154">SyncFolderItems operation</span></span>](syncfolderitems-operation.md)
- [<span data-ttu-id="3638c-155">Exchange の EWS XML 要素</span><span class="sxs-lookup"><span data-stu-id="3638c-155">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)

