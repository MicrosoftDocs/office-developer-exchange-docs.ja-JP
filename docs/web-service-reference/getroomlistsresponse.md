---
title: GetRoomListsResponse
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- GetRoomListsResponse
api_type:
- schema
ms.assetid: 8c736f68-1026-496a-b12f-c169c078abd0
description: GetRoomListsResponse 要素は、GetRoomLists 操作の要求からの応答を定義します。
ms.openlocfilehash: 8231435f7dc348a070852f57d6152550326b5df6
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19760855"
---
# <a name="getroomlistsresponse"></a><span data-ttu-id="ee34c-103">GetRoomListsResponse</span><span class="sxs-lookup"><span data-stu-id="ee34c-103">GetRoomListsResponse</span></span>

<span data-ttu-id="ee34c-104">**GetRoomListsResponse**要素は、 [GetRoomLists 操作](getroomlists-operation.md)の要求からの応答を定義します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-104">The **GetRoomListsResponse** element defines the response from a [GetRoomLists operation](getroomlists-operation.md) request.</span></span> 
  
- [<span data-ttu-id="ee34c-105">ResponseMessages</span><span class="sxs-lookup"><span data-stu-id="ee34c-105">ResponseMessages</span></span>](responsemessages.md)
- [<span data-ttu-id="ee34c-106">GetRoomListsResponse</span><span class="sxs-lookup"><span data-stu-id="ee34c-106">GetRoomListsResponse</span></span>](getroomlistsresponse.md)
  
```XML
<GetRoomListsResponse ResponseClass="">   
<MessageText/>   
<ResponseCode/>   
<DescriptiveLinkKey/>   
<MessageXml/>   
<RoomLists/>
</GetRoomListsResponse>
```

 <span data-ttu-id="ee34c-107">**GetRoomListsResponseMessageType**</span><span class="sxs-lookup"><span data-stu-id="ee34c-107">**GetRoomListsResponseMessageType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="ee34c-108">属性および要素</span><span class="sxs-lookup"><span data-stu-id="ee34c-108">Attributes and elements</span></span>

<span data-ttu-id="ee34c-109">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-109">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="ee34c-110">属性</span><span class="sxs-lookup"><span data-stu-id="ee34c-110">Attributes</span></span>

|<span data-ttu-id="ee34c-111">**属性**</span><span class="sxs-lookup"><span data-stu-id="ee34c-111">**Attribute**</span></span>|<span data-ttu-id="ee34c-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="ee34c-112">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ee34c-113">**ResponseClass**</span><span class="sxs-lookup"><span data-stu-id="ee34c-113">**ResponseClass**</span></span> <br/> | <span data-ttu-id="ee34c-114">応答のステータスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-114">Describes the status of the response.</span></span> <br/><br/><span data-ttu-id="ee34c-115">次の値は、この属性の有効です。</span><span class="sxs-lookup"><span data-stu-id="ee34c-115">The following values are valid for this attribute:</span></span>  <br/><br/><span data-ttu-id="ee34c-116">-成功</span><span class="sxs-lookup"><span data-stu-id="ee34c-116">-  Success</span></span>  <br/><span data-ttu-id="ee34c-117">-警告</span><span class="sxs-lookup"><span data-stu-id="ee34c-117">-  Warning</span></span>  <br/><span data-ttu-id="ee34c-118">-エラー</span><span class="sxs-lookup"><span data-stu-id="ee34c-118">-  Error</span></span>  <br/> |
   
#### <a name="responseclass-attribute-values"></a><span data-ttu-id="ee34c-119">ResponseClass 属性の値</span><span class="sxs-lookup"><span data-stu-id="ee34c-119">ResponseClass attribute values</span></span>

|<span data-ttu-id="ee34c-120">**値**</span><span class="sxs-lookup"><span data-stu-id="ee34c-120">**Value**</span></span>|<span data-ttu-id="ee34c-121">**説明**</span><span class="sxs-lookup"><span data-stu-id="ee34c-121">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ee34c-122">**Success**</span><span class="sxs-lookup"><span data-stu-id="ee34c-122">**Success**</span></span> <br/> |<span data-ttu-id="ee34c-123">満了する要求をについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-123">Describes a request that is fulfilled.</span></span>  <br/> |
|<span data-ttu-id="ee34c-124">**Warning**</span><span class="sxs-lookup"><span data-stu-id="ee34c-124">**Warning**</span></span> <br/> | <span data-ttu-id="ee34c-125">処理されなかった要求をについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-125">Describes a request that was not processed.</span></span> <span data-ttu-id="ee34c-126">警告は、要求内のアイテムを処理し、それ以降の項目を処理できませんでした、エラーが発生した場合に返される場合があります。</span><span class="sxs-lookup"><span data-stu-id="ee34c-126">A warning may be returned if an error occurred while an item in the request was processing and subsequent items could not be processed.</span></span> <br/><br/><span data-ttu-id="ee34c-127">警告の送信元の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-127">The following are examples of sources of warnings:</span></span>  <br/><br/><span data-ttu-id="ee34c-128">-Exchange ストアは、バッチの実行中にオフラインです。</span><span class="sxs-lookup"><span data-stu-id="ee34c-128">-  The Exchange store is offline during the batch.</span></span>  <br/><span data-ttu-id="ee34c-129">-Active Directory ディレクトリ サービスは、オフラインです。</span><span class="sxs-lookup"><span data-stu-id="ee34c-129">-  The Active Directory directory service is offline.</span></span>  <br/><span data-ttu-id="ee34c-130">-メールボックスは移動されました。</span><span class="sxs-lookup"><span data-stu-id="ee34c-130">-  Mailboxes were moved.</span></span>  <br/><span data-ttu-id="ee34c-131">-メッセージ データベース (MDB) は、オフラインです。</span><span class="sxs-lookup"><span data-stu-id="ee34c-131">-  The message database (MDB) is offline.</span></span>  <br/><span data-ttu-id="ee34c-132">-パスワードの有効期限が切れています。</span><span class="sxs-lookup"><span data-stu-id="ee34c-132">-  A password is expired.</span></span>  <br/><span data-ttu-id="ee34c-133">クォータを超えています。</span><span class="sxs-lookup"><span data-stu-id="ee34c-133">-  A quota has been exceeded.</span></span>  <br/> |
|<span data-ttu-id="ee34c-134">**Error**</span><span class="sxs-lookup"><span data-stu-id="ee34c-134">**Error**</span></span> <br/> | <span data-ttu-id="ee34c-135">満たせない要求をについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-135">Describes a request that cannot be fulfilled.</span></span> <br/><br/><span data-ttu-id="ee34c-136">次に、エラーのソースの例を示します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-136">The following are examples of sources of errors:</span></span>  <br/><br/><span data-ttu-id="ee34c-137">-無効な属性または要素</span><span class="sxs-lookup"><span data-stu-id="ee34c-137">-  Invalid attributes or elements</span></span>  <br/><span data-ttu-id="ee34c-138">属性または要素が範囲外です</span><span class="sxs-lookup"><span data-stu-id="ee34c-138">-  Attributes or elements out of range</span></span>  <br/><span data-ttu-id="ee34c-139">-不明なタグ</span><span class="sxs-lookup"><span data-stu-id="ee34c-139">-  Unknown tag</span></span>  <br/><span data-ttu-id="ee34c-140">属性または要素のコンテキストでは無効です</span><span class="sxs-lookup"><span data-stu-id="ee34c-140">-  Attribute or element not valid in the context</span></span>  <br/><span data-ttu-id="ee34c-141">の任意のクライアントから不正アクセスしようと</span><span class="sxs-lookup"><span data-stu-id="ee34c-141">-  Unauthorized access attempt by any client</span></span>  <br/><span data-ttu-id="ee34c-142">の有効なクライアント側の呼び出しに応答サーバー側の障害</span><span class="sxs-lookup"><span data-stu-id="ee34c-142">-  Server-side failure in response to a valid client-side call</span></span>  <br/><br/>  <span data-ttu-id="ee34c-143">エラーに関する情報は、 [ResponseCode](responsecode.md)および[メッセージ テキスト](messagetext.md)の要素を参照しています。</span><span class="sxs-lookup"><span data-stu-id="ee34c-143">Information about the error can be found in the [ResponseCode](responsecode.md) and [MessageText](messagetext.md) elements.</span></span>  <br/>- |
   
### <a name="child-elements"></a><span data-ttu-id="ee34c-144">子要素</span><span class="sxs-lookup"><span data-stu-id="ee34c-144">Child elements</span></span>

|<span data-ttu-id="ee34c-145">**要素**</span><span class="sxs-lookup"><span data-stu-id="ee34c-145">**Element**</span></span>|<span data-ttu-id="ee34c-146">**説明**</span><span class="sxs-lookup"><span data-stu-id="ee34c-146">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="ee34c-147">MessageText</span><span class="sxs-lookup"><span data-stu-id="ee34c-147">MessageText</span></span>](messagetext.md) <br/> |<span data-ttu-id="ee34c-148">応答のステータスの説明を提供します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-148">Provides a text description of the status of the response.</span></span>  <br/> |
|[<span data-ttu-id="ee34c-149">ResponseCode</span><span class="sxs-lookup"><span data-stu-id="ee34c-149">ResponseCode</span></span>](responsecode.md) <br/> |<span data-ttu-id="ee34c-150">要求で発生した特定のエラーを識別するエラー コードを提供します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-150">Provides an error code that identifies the specific error that the request encountered.</span></span>  <br/> |
|[<span data-ttu-id="ee34c-151">DescriptiveLinkKey</span><span class="sxs-lookup"><span data-stu-id="ee34c-151">DescriptiveLinkKey</span></span>](descriptivelinkkey.md) <br/> |<span data-ttu-id="ee34c-152">現在使用されていない、将来使用するために予約されています。</span><span class="sxs-lookup"><span data-stu-id="ee34c-152">Currently unused and reserved for future use.</span></span> <span data-ttu-id="ee34c-153">この要素には、0 の値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ee34c-153">This element contains a value of 0.</span></span>  <br/> |
|[<span data-ttu-id="ee34c-154">MessageXml</span><span class="sxs-lookup"><span data-stu-id="ee34c-154">MessageXml</span></span>](messagexml.md) <br/> |<span data-ttu-id="ee34c-155">追加のエラー応答情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-155">Provides additional error response information.</span></span>  <br/> |
|[<span data-ttu-id="ee34c-156">RoomLists</span><span class="sxs-lookup"><span data-stu-id="ee34c-156">RoomLists</span></span>](roomlists.md) <br/> |<span data-ttu-id="ee34c-157">電子メール アドレスや会議室の一覧を表す表示名の一覧を示します。</span><span class="sxs-lookup"><span data-stu-id="ee34c-157">Provides a list of e-mail addresses and display names that represent lists of meeting rooms.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="ee34c-158">親要素</span><span class="sxs-lookup"><span data-stu-id="ee34c-158">Parent elements</span></span>

|<span data-ttu-id="ee34c-159">**要素**</span><span class="sxs-lookup"><span data-stu-id="ee34c-159">**Element**</span></span>|<span data-ttu-id="ee34c-160">**説明**</span><span class="sxs-lookup"><span data-stu-id="ee34c-160">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="ee34c-161">ResponseMessages</span><span class="sxs-lookup"><span data-stu-id="ee34c-161">ResponseMessages</span></span>](responsemessages.md) <br/> |<span data-ttu-id="ee34c-162">Exchange Web サービス要求に対する応答メッセージが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ee34c-162">Contains the response messages for an Exchange Web Services request.</span></span>  <br/> |
   
## <a name="element-information"></a><span data-ttu-id="ee34c-163">要素情報</span><span class="sxs-lookup"><span data-stu-id="ee34c-163">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="ee34c-164">名前空間</span><span class="sxs-lookup"><span data-stu-id="ee34c-164">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="ee34c-165">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="ee34c-165">Schema Name</span></span>  <br/> |<span data-ttu-id="ee34c-166">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="ee34c-166">Messages schema</span></span>  <br/> |
|<span data-ttu-id="ee34c-167">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="ee34c-167">Validation File</span></span>  <br/> |<span data-ttu-id="ee34c-168">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="ee34c-168">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="ee34c-169">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="ee34c-169">Can be Empty</span></span>  <br/> |<span data-ttu-id="ee34c-170">False</span><span class="sxs-lookup"><span data-stu-id="ee34c-170">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="ee34c-171">関連項目</span><span class="sxs-lookup"><span data-stu-id="ee34c-171">See also</span></span>

- [<span data-ttu-id="ee34c-172">GetRoomLists 操作</span><span class="sxs-lookup"><span data-stu-id="ee34c-172">GetRoomLists operation</span></span>](getroomlists-operation.md)
- [<span data-ttu-id="ee34c-173">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="ee34c-173">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
