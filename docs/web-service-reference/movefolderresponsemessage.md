---
title: MoveFolderResponseMessage
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MoveFolderResponseMessage
api_type:
- schema
ms.assetid: 54917251-96af-44c2-ae90-d545f0a16e2e
description: MoveFolderResponseMessage 要素には、状態および 1 つの MoveFolder 操作要求の結果が含まれています。
ms.openlocfilehash: 777520bf6a093882da95a7bfd466b66acdab957c
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832496"
---
# <a name="movefolderresponsemessage"></a><span data-ttu-id="4eab0-103">MoveFolderResponseMessage</span><span class="sxs-lookup"><span data-stu-id="4eab0-103">MoveFolderResponseMessage</span></span>

<span data-ttu-id="4eab0-104">**MoveFolderResponseMessage**要素には、状態および 1 つの結果が含まれています[MoveFolder 操作](movefolder-operation.md)を要求します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-104">The **MoveFolderResponseMessage** element contains the status and result of a single [MoveFolder operation](movefolder-operation.md) request.</span></span> 
  
- [<span data-ttu-id="4eab0-105">MoveFolderResponse</span><span class="sxs-lookup"><span data-stu-id="4eab0-105">MoveFolderResponse</span></span>](movefolderresponse.md)
- [<span data-ttu-id="4eab0-106">ResponseMessages</span><span class="sxs-lookup"><span data-stu-id="4eab0-106">ResponseMessages</span></span>](responsemessages.md)
- [<span data-ttu-id="4eab0-107">MoveFolderResponseMessage</span><span class="sxs-lookup"><span data-stu-id="4eab0-107">MoveFolderResponseMessage</span></span>](movefolderresponsemessage.md)
  
```xml
<MoveFolderResponseMessage ResponseClass="">
   <MessageText/>
   <ResponseCode/>
   <DescriptiveLinkKey/>
   <MessageXml/>
   <Folders/>
</MoveFolderResponseMessage>
```

 <span data-ttu-id="4eab0-108">**FolderInfoResponseMessageType**</span><span class="sxs-lookup"><span data-stu-id="4eab0-108">**FolderInfoResponseMessageType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="4eab0-109">属性および要素</span><span class="sxs-lookup"><span data-stu-id="4eab0-109">Attributes and elements</span></span>

<span data-ttu-id="4eab0-110">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-110">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="4eab0-111">属性</span><span class="sxs-lookup"><span data-stu-id="4eab0-111">Attributes</span></span>

|<span data-ttu-id="4eab0-112">**属性**</span><span class="sxs-lookup"><span data-stu-id="4eab0-112">**Attribute**</span></span>|<span data-ttu-id="4eab0-113">**説明**</span><span class="sxs-lookup"><span data-stu-id="4eab0-113">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="4eab0-114">**ResponseClass**</span><span class="sxs-lookup"><span data-stu-id="4eab0-114">**ResponseClass**</span></span> <br/> | <span data-ttu-id="4eab0-115">[MoveFolder 操作](movefolder-operation.md)の応答のステータスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-115">Describes the status of a [MoveFolder operation](movefolder-operation.md) response.</span></span> <br/><br/><span data-ttu-id="4eab0-116">次の値は、この属性の有効です。</span><span class="sxs-lookup"><span data-stu-id="4eab0-116">The following values are valid for this attribute:</span></span>  <br/><br/><span data-ttu-id="4eab0-117">-成功</span><span class="sxs-lookup"><span data-stu-id="4eab0-117">-  Success</span></span>  <br/><span data-ttu-id="4eab0-118">-警告</span><span class="sxs-lookup"><span data-stu-id="4eab0-118">-  Warning</span></span>  <br/><span data-ttu-id="4eab0-119">-エラー</span><span class="sxs-lookup"><span data-stu-id="4eab0-119">-  Error</span></span>  <br/> |
   
#### <a name="responseclass-attribute"></a><span data-ttu-id="4eab0-120">ResponseClass 属性</span><span class="sxs-lookup"><span data-stu-id="4eab0-120">ResponseClass Attribute</span></span>

|<span data-ttu-id="4eab0-121">**値**</span><span class="sxs-lookup"><span data-stu-id="4eab0-121">**Value**</span></span>|<span data-ttu-id="4eab0-122">**説明**</span><span class="sxs-lookup"><span data-stu-id="4eab0-122">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="4eab0-123">**Success**</span><span class="sxs-lookup"><span data-stu-id="4eab0-123">**Success**</span></span> <br/> |<span data-ttu-id="4eab0-124">満了する要求をについて説明します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-124">Describes a request that is fulfilled.</span></span>  <br/> |
|<span data-ttu-id="4eab0-125">**Warning**</span><span class="sxs-lookup"><span data-stu-id="4eab0-125">**Warning**</span></span> <br/> | <span data-ttu-id="4eab0-126">処理されなかった要求をについて説明します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-126">Describes a request that was not processed.</span></span> <span data-ttu-id="4eab0-127">警告は、要求内のアイテムを処理し、それ以降の項目を処理できませんでした、エラーが発生した場合に返される場合があります。</span><span class="sxs-lookup"><span data-stu-id="4eab0-127">A warning may be returned if an error occurred while an item in the request was processing and subsequent items could not be processed.</span></span><br/><br/> <span data-ttu-id="4eab0-128">警告の送信元の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-128">The following are examples of sources of warnings:</span></span>  <br/><br/><span data-ttu-id="4eab0-129">-Exchange ストアは、バッチの実行中にオフラインです。</span><span class="sxs-lookup"><span data-stu-id="4eab0-129">-  The Exchange store is offline during the batch.</span></span>  <br/><span data-ttu-id="4eab0-130">-Active Directory ドメイン サービス (AD DS) では、オフラインです。</span><span class="sxs-lookup"><span data-stu-id="4eab0-130">-  Active Directory Domain Services (AD DS) is offline.</span></span>  <br/><span data-ttu-id="4eab0-131">-メールボックスは移動されました。</span><span class="sxs-lookup"><span data-stu-id="4eab0-131">-  Mailboxes were moved.</span></span>  <br/><span data-ttu-id="4eab0-132">-メッセージ データベース (MDB) は、オフラインです。</span><span class="sxs-lookup"><span data-stu-id="4eab0-132">-  The message database (MDB) is offline.</span></span>  <br/><span data-ttu-id="4eab0-133">-パスワードの有効期限が切れています。</span><span class="sxs-lookup"><span data-stu-id="4eab0-133">-  A password is expired.</span></span>  <br/><span data-ttu-id="4eab0-134">クォータを超過しました。</span><span class="sxs-lookup"><span data-stu-id="4eab0-134">-  A quota was exceeded.</span></span>  <br/> |
|<span data-ttu-id="4eab0-135">**Error**</span><span class="sxs-lookup"><span data-stu-id="4eab0-135">**Error**</span></span> <br/> | <span data-ttu-id="4eab0-136">満たせない要求をについて説明します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-136">Describes a request that cannot be fulfilled.</span></span> <br/><br/><span data-ttu-id="4eab0-137">次に、エラーのソースの例を示します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-137">The following are examples of sources of errors:</span></span>  <br/><br/><span data-ttu-id="4eab0-138">-無効な属性または要素</span><span class="sxs-lookup"><span data-stu-id="4eab0-138">-  Invalid attributes or elements</span></span>  <br/><span data-ttu-id="4eab0-139">属性または要素が範囲外です</span><span class="sxs-lookup"><span data-stu-id="4eab0-139">-  Attributes or elements out of range</span></span>  <br/><span data-ttu-id="4eab0-140">-不明なタグ</span><span class="sxs-lookup"><span data-stu-id="4eab0-140">-  Unknown tag</span></span>  <br/><span data-ttu-id="4eab0-141">属性または要素のコンテキストでは無効です</span><span class="sxs-lookup"><span data-stu-id="4eab0-141">-  Attribute or element not valid in the context</span></span>  <br/><span data-ttu-id="4eab0-142">-任意のクライアントから許可されていないアクセス試行</span><span class="sxs-lookup"><span data-stu-id="4eab0-142">-  Any unauthorized access attempt by any client</span></span>  <br/><span data-ttu-id="4eab0-143">-有効なクライアント側の呼び出しへの応答でサーバー側の障害</span><span class="sxs-lookup"><span data-stu-id="4eab0-143">-  Any server-side failure in response to a valid client-side call</span></span>  <br/><br/>  <span data-ttu-id="4eab0-144">エラーに関する情報は、 [ResponseCode](responsecode.md)および[メッセージ テキスト](messagetext.md)の要素を参照しています。</span><span class="sxs-lookup"><span data-stu-id="4eab0-144">Information about the error can be found in the [ResponseCode](responsecode.md) and [MessageText](messagetext.md) elements.</span></span>  <br/> |
   
### <a name="child-elements"></a><span data-ttu-id="4eab0-145">子要素</span><span class="sxs-lookup"><span data-stu-id="4eab0-145">Child elements</span></span>

|<span data-ttu-id="4eab0-146">**要素**</span><span class="sxs-lookup"><span data-stu-id="4eab0-146">**Element**</span></span>|<span data-ttu-id="4eab0-147">**説明**</span><span class="sxs-lookup"><span data-stu-id="4eab0-147">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="4eab0-148">MessageText</span><span class="sxs-lookup"><span data-stu-id="4eab0-148">MessageText</span></span>](messagetext.md) <br/> |<span data-ttu-id="4eab0-149">応答のステータスの説明です。</span><span class="sxs-lookup"><span data-stu-id="4eab0-149">A text description of the status of the response.</span></span>  <br/> |
|[<span data-ttu-id="4eab0-150">ResponseCode</span><span class="sxs-lookup"><span data-stu-id="4eab0-150">ResponseCode</span></span>](responsecode.md) <br/> |<span data-ttu-id="4eab0-151">要求で発生した特定のエラーを識別するエラー コードを提供します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-151">Provides an error code that identifies the specific error that the request encountered.</span></span>  <br/> |
|[<span data-ttu-id="4eab0-152">DescriptiveLinkKey</span><span class="sxs-lookup"><span data-stu-id="4eab0-152">DescriptiveLinkKey</span></span>](descriptivelinkkey.md) <br/> |<span data-ttu-id="4eab0-153">現在使用されていない将来の使用に予約されているとします。</span><span class="sxs-lookup"><span data-stu-id="4eab0-153">Currently unused and is reserved for future use.</span></span> <span data-ttu-id="4eab0-154">0 の値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4eab0-154">It contains a value of 0.</span></span>  <br/> |
|[<span data-ttu-id="4eab0-155">MessageXml</span><span class="sxs-lookup"><span data-stu-id="4eab0-155">MessageXml</span></span>](messagexml.md) <br/> |<span data-ttu-id="4eab0-156">追加のエラー応答情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="4eab0-156">Provides additional error response information.</span></span>  <br/> |
|[<span data-ttu-id="4eab0-157">フォルダー</span><span class="sxs-lookup"><span data-stu-id="4eab0-157">Folders</span></span>](folders-ex15websvcsotherref.md) <br/> |<span data-ttu-id="4eab0-158">移動されたフォルダーの配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4eab0-158">Contains an array of moved folders.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="4eab0-159">親要素</span><span class="sxs-lookup"><span data-stu-id="4eab0-159">Parent elements</span></span>

|<span data-ttu-id="4eab0-160">**要素**</span><span class="sxs-lookup"><span data-stu-id="4eab0-160">**Element**</span></span>|<span data-ttu-id="4eab0-161">**説明**</span><span class="sxs-lookup"><span data-stu-id="4eab0-161">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="4eab0-162">ResponseMessages</span><span class="sxs-lookup"><span data-stu-id="4eab0-162">ResponseMessages</span></span>](responsemessages.md) <br/> |<span data-ttu-id="4eab0-163">Exchange Web サービス要求に対する応答メッセージが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4eab0-163">Contains the response messages for an Exchange Web Services request.</span></span>  <br/> |
   
## <a name="remarks"></a><span data-ttu-id="4eab0-164">備考</span><span class="sxs-lookup"><span data-stu-id="4eab0-164">Remarks</span></span>

<span data-ttu-id="4eab0-165">この要素を記述するスキーマは、クライアント アクセス サーバーの役割がインストールされている Microsoft Exchange Server 2010 を実行しているコンピューターの EWS 仮想ディレクトリにあります。</span><span class="sxs-lookup"><span data-stu-id="4eab0-165">The schema that describes this element is located in the EWS virtual directory of the computer that is running Microsoft Exchange Server 2010 that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="4eab0-166">要素情報</span><span class="sxs-lookup"><span data-stu-id="4eab0-166">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="4eab0-167">名前空間</span><span class="sxs-lookup"><span data-stu-id="4eab0-167">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="4eab0-168">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="4eab0-168">Schema Name</span></span>  <br/> |<span data-ttu-id="4eab0-169">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="4eab0-169">Messages schema</span></span>  <br/> |
|<span data-ttu-id="4eab0-170">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="4eab0-170">Validation File</span></span>  <br/> |<span data-ttu-id="4eab0-171">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="4eab0-171">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="4eab0-172">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="4eab0-172">Can be Empty</span></span>  <br/> |<span data-ttu-id="4eab0-173">False</span><span class="sxs-lookup"><span data-stu-id="4eab0-173">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="4eab0-174">関連項目</span><span class="sxs-lookup"><span data-stu-id="4eab0-174">See also</span></span>

- [<span data-ttu-id="4eab0-175">MoveFolder</span><span class="sxs-lookup"><span data-stu-id="4eab0-175">MoveFolder</span></span>](movefolder.md)
- [<span data-ttu-id="4eab0-176">MoveFolder 操作</span><span class="sxs-lookup"><span data-stu-id="4eab0-176">MoveFolder operation</span></span>](movefolder-operation.md)
