---
title: RoomList
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- RoomList
api_type:
- schema
ms.assetid: cb02bdf0-df9f-4e31-b7dd-cd9f2f2cc2b2
description: RoomList 要素は、会議室の一覧を識別する電子メール アドレスを表します。
ms.openlocfilehash: 7de2c67f8001387abf463186933f0b81ee45a58a
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19833262"
---
# <a name="roomlist"></a><span data-ttu-id="2f266-103">RoomList</span><span class="sxs-lookup"><span data-stu-id="2f266-103">RoomList</span></span>

<span data-ttu-id="2f266-104">**RoomList**要素は、会議室の一覧を識別する電子メール アドレスを表します。</span><span class="sxs-lookup"><span data-stu-id="2f266-104">The **RoomList** element represents an e-mail address that identifies a list of meeting rooms.</span></span> 
  
[<span data-ttu-id="2f266-105">GetRooms</span><span class="sxs-lookup"><span data-stu-id="2f266-105">GetRooms</span></span>](getrooms.md)
  
[<span data-ttu-id="2f266-106">RoomList</span><span class="sxs-lookup"><span data-stu-id="2f266-106">RoomList</span></span>](roomlist.md)
  
```XML
<RoomList>
   <Name/>
   <EmailAddress/>
   <RoutingType/>
   <MailboxType/>
   <ItemId/>
</RoomList>
```

 <span data-ttu-id="2f266-107">**EmailAddressType**</span><span class="sxs-lookup"><span data-stu-id="2f266-107">**EmailAddressType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="2f266-108">属性および要素</span><span class="sxs-lookup"><span data-stu-id="2f266-108">Attributes and elements</span></span>

<span data-ttu-id="2f266-109">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="2f266-109">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="2f266-110">属性</span><span class="sxs-lookup"><span data-stu-id="2f266-110">Attributes</span></span>

<span data-ttu-id="2f266-111">なし。</span><span class="sxs-lookup"><span data-stu-id="2f266-111">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="2f266-112">子要素</span><span class="sxs-lookup"><span data-stu-id="2f266-112">Child elements</span></span>

|<span data-ttu-id="2f266-113">**要素**</span><span class="sxs-lookup"><span data-stu-id="2f266-113">**Element**</span></span>|<span data-ttu-id="2f266-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="2f266-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="2f266-115">名 (EmailAddressType)</span><span class="sxs-lookup"><span data-stu-id="2f266-115">Name (EmailAddressType)</span></span>](name-emailaddresstype.md) <br/> |<span data-ttu-id="2f266-116">ルーム リストの表示名を定義します。</span><span class="sxs-lookup"><span data-stu-id="2f266-116">Defines the display name of the room list.</span></span> <span data-ttu-id="2f266-117">この要素はオプションです。</span><span class="sxs-lookup"><span data-stu-id="2f266-117">This element is optional.</span></span>  <br/> |
|[<span data-ttu-id="2f266-118">EmailAddress (NonEmptyStringType)</span><span class="sxs-lookup"><span data-stu-id="2f266-118">EmailAddress (NonEmptyStringType)</span></span>](emailaddress-nonemptystringtype.md) <br/> |<span data-ttu-id="2f266-119">会議室の一覧の簡易メール転送プロトコル (SMTP) アドレスを定義します。</span><span class="sxs-lookup"><span data-stu-id="2f266-119">Defines the Simple Mail Transfer Protocol (SMTP) address of a room list.</span></span> <span data-ttu-id="2f266-120">この要素はオプションです。</span><span class="sxs-lookup"><span data-stu-id="2f266-120">This element is optional.</span></span>  <br/> |
|[<span data-ttu-id="2f266-121">RoutingType (EmailAddress)</span><span class="sxs-lookup"><span data-stu-id="2f266-121">RoutingType (EmailAddress)</span></span>](routingtype-emailaddress.md) <br/> |<span data-ttu-id="2f266-122">メールボックスに使用されるルーティングを定義します。</span><span class="sxs-lookup"><span data-stu-id="2f266-122">Defines the routing that is used for the mailbox.</span></span> <span data-ttu-id="2f266-123">既定値は、SMTP です。</span><span class="sxs-lookup"><span data-stu-id="2f266-123">The default is SMTP.</span></span> <span data-ttu-id="2f266-124">この要素はオプションです。</span><span class="sxs-lookup"><span data-stu-id="2f266-124">This element is optional.</span></span>  <br/> |
|[<span data-ttu-id="2f266-125">MailboxType</span><span class="sxs-lookup"><span data-stu-id="2f266-125">MailboxType</span></span>](mailboxtype.md) <br/> |<span data-ttu-id="2f266-126">メールボックスのユーザーのメールボックスの種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="2f266-126">Defines the mailbox type of a mailbox user.</span></span> <span data-ttu-id="2f266-127">この要素はオプションです。</span><span class="sxs-lookup"><span data-stu-id="2f266-127">This element is optional.</span></span>  <br/> |
|[<span data-ttu-id="2f266-128">ItemId</span><span class="sxs-lookup"><span data-stu-id="2f266-128">ItemId</span></span>](itemid.md) <br/> |<span data-ttu-id="2f266-129">ユーザーのアドレス帳フォルダーから受信者の連絡先または個人用配布リストの項目の識別子を定義します。</span><span class="sxs-lookup"><span data-stu-id="2f266-129">Defines the item identifier of a contact or private distribution list for recipients from a user's Contacts folder.</span></span> <span data-ttu-id="2f266-130">この要素はオプションです。</span><span class="sxs-lookup"><span data-stu-id="2f266-130">This element is optional.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="2f266-131">親要素</span><span class="sxs-lookup"><span data-stu-id="2f266-131">Parent elements</span></span>

|<span data-ttu-id="2f266-132">**要素**</span><span class="sxs-lookup"><span data-stu-id="2f266-132">**Element**</span></span>|<span data-ttu-id="2f266-133">**説明**</span><span class="sxs-lookup"><span data-stu-id="2f266-133">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="2f266-134">GetRooms</span><span class="sxs-lookup"><span data-stu-id="2f266-134">GetRooms</span></span>](getrooms.md) <br/> |<span data-ttu-id="2f266-135">特定の会議室の一覧内の会議室の一覧を取得する要求のルート要素です。</span><span class="sxs-lookup"><span data-stu-id="2f266-135">The root element in a request to get a list of rooms within a particular room list.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="2f266-136">テキスト値</span><span class="sxs-lookup"><span data-stu-id="2f266-136">Text value</span></span>

<span data-ttu-id="2f266-137">なし。</span><span class="sxs-lookup"><span data-stu-id="2f266-137">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="2f266-138">備考</span><span class="sxs-lookup"><span data-stu-id="2f266-138">Remarks</span></span>

<span data-ttu-id="2f266-139">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="2f266-139">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="2f266-140">要素情報</span><span class="sxs-lookup"><span data-stu-id="2f266-140">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="2f266-141">名前空間</span><span class="sxs-lookup"><span data-stu-id="2f266-141">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="2f266-142">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="2f266-142">Schema Name</span></span>  <br/> |<span data-ttu-id="2f266-143">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="2f266-143">Messages schema</span></span>  <br/> |
|<span data-ttu-id="2f266-144">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="2f266-144">Validation File</span></span>  <br/> |<span data-ttu-id="2f266-145">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="2f266-145">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="2f266-146">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="2f266-146">Can be Empty</span></span>  <br/> |<span data-ttu-id="2f266-147">False</span><span class="sxs-lookup"><span data-stu-id="2f266-147">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="2f266-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="2f266-148">See also</span></span>



[<span data-ttu-id="2f266-149">GetRooms 操作</span><span class="sxs-lookup"><span data-stu-id="2f266-149">GetRooms operation</span></span>](getrooms-operation.md)


- [<span data-ttu-id="2f266-150">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="2f266-150">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
