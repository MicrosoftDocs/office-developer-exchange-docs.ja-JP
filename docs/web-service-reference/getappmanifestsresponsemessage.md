---
title: GetAppManifestsResponseMessage
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 815908f1-4223-42d8-92dc-f8bdfc6b5df8
description: GetAppManifestsResponseMessage 要素は、GetAppManifests 要求の応答メッセージを指定します。
ms.openlocfilehash: 05eeef7f7194c1dc05be93ed13ebff93d5013e94
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19760632"
---
# <a name="getappmanifestsresponsemessage"></a><span data-ttu-id="48f7b-103">GetAppManifestsResponseMessage</span><span class="sxs-lookup"><span data-stu-id="48f7b-103">GetAppManifestsResponseMessage</span></span>

<span data-ttu-id="48f7b-104">**GetAppManifestsResponseMessage**要素は、 **GetAppManifests**要求の応答メッセージを指定します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-104">The **GetAppManifestsResponseMessage** element specifies the response message for a **GetAppManifests** request.</span></span> 
  
```XML
<GetAppManifestsResponseMessage ResponseClass=" Success | Warning | Error ">
    <MessageText/>
    <ResponseCode/>
    <DescriptiveLinkKey/>
    <MessageXml/>
</GetAppManifestsResponseMessage>
```

 <span data-ttu-id="48f7b-105">**ResponseMessageType**</span><span class="sxs-lookup"><span data-stu-id="48f7b-105">**ResponseMessageType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="48f7b-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="48f7b-106">Attributes and elements</span></span>

<span data-ttu-id="48f7b-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="48f7b-108">属性</span><span class="sxs-lookup"><span data-stu-id="48f7b-108">Attributes</span></span>

|<span data-ttu-id="48f7b-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="48f7b-109">**Attribute**</span></span>|<span data-ttu-id="48f7b-110">**説明**</span><span class="sxs-lookup"><span data-stu-id="48f7b-110">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="48f7b-111">ResponseClass</span><span class="sxs-lookup"><span data-stu-id="48f7b-111">ResponseClass</span></span>  <br/> |<span data-ttu-id="48f7b-112">応答のクラスを示します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-112">Indicates the class of the response.</span></span>  <br/> |
   
#### <a name="responseclass"></a><span data-ttu-id="48f7b-113">ResponseClass</span><span class="sxs-lookup"><span data-stu-id="48f7b-113">ResponseClass</span></span>

|<span data-ttu-id="48f7b-114">**値**</span><span class="sxs-lookup"><span data-stu-id="48f7b-114">**Value**</span></span>|<span data-ttu-id="48f7b-115">**説明**</span><span class="sxs-lookup"><span data-stu-id="48f7b-115">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="48f7b-116">成功</span><span class="sxs-lookup"><span data-stu-id="48f7b-116">Success</span></span>  <br/> |<span data-ttu-id="48f7b-117">成功を示します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-117">Indicates success.</span></span>  <br/> |
|<span data-ttu-id="48f7b-118">警告</span><span class="sxs-lookup"><span data-stu-id="48f7b-118">Warning</span></span>  <br/> |<span data-ttu-id="48f7b-119">警告を示します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-119">Indicates a warning.</span></span>  <br/> |
|<span data-ttu-id="48f7b-120">エラー</span><span class="sxs-lookup"><span data-stu-id="48f7b-120">Error</span></span>  <br/> |<span data-ttu-id="48f7b-121">エラーを示します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-121">Indicates an error.</span></span>  <br/> |
   
### <a name="child-elements"></a><span data-ttu-id="48f7b-122">子要素</span><span class="sxs-lookup"><span data-stu-id="48f7b-122">Child elements</span></span>

|<span data-ttu-id="48f7b-123">**要素**</span><span class="sxs-lookup"><span data-stu-id="48f7b-123">**Element**</span></span>|<span data-ttu-id="48f7b-124">**説明**</span><span class="sxs-lookup"><span data-stu-id="48f7b-124">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="48f7b-125">DescriptiveLinkKey</span><span class="sxs-lookup"><span data-stu-id="48f7b-125">DescriptiveLinkKey</span></span>](descriptivelinkkey.md) <br/> |<span data-ttu-id="48f7b-126">現在使用されていない、将来使用するために予約されています。</span><span class="sxs-lookup"><span data-stu-id="48f7b-126">Currently unused and reserved for future use.</span></span>  <br/> |
|[<span data-ttu-id="48f7b-127">MessageText</span><span class="sxs-lookup"><span data-stu-id="48f7b-127">MessageText</span></span>](messagetext.md) <br/> |<span data-ttu-id="48f7b-128">応答のステータスの説明を提供します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-128">Provides a text description of the status of the response.</span></span>  <br/> |
|[<span data-ttu-id="48f7b-129">MessageXml</span><span class="sxs-lookup"><span data-stu-id="48f7b-129">MessageXml</span></span>](messagexml.md) <br/> |<span data-ttu-id="48f7b-130">追加のエラー応答情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-130">Provides additional error response information.</span></span>  <br/> |
|[<span data-ttu-id="48f7b-131">ResponseCode</span><span class="sxs-lookup"><span data-stu-id="48f7b-131">ResponseCode</span></span>](responsecode.md) <br/> |<span data-ttu-id="48f7b-132">要求に関するステータス情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="48f7b-132">Provides status information about the request.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="48f7b-133">親要素</span><span class="sxs-lookup"><span data-stu-id="48f7b-133">Parent elements</span></span>

|<span data-ttu-id="48f7b-134">**要素**</span><span class="sxs-lookup"><span data-stu-id="48f7b-134">**Element**</span></span>|<span data-ttu-id="48f7b-135">**説明**</span><span class="sxs-lookup"><span data-stu-id="48f7b-135">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="48f7b-136">ResponseMessages</span><span class="sxs-lookup"><span data-stu-id="48f7b-136">ResponseMessages</span></span>](responsemessages.md) <br/> |<span data-ttu-id="48f7b-137">Exchange Web サービス要求に対する応答メッセージが含まれています。</span><span class="sxs-lookup"><span data-stu-id="48f7b-137">Contains the response messages for an Exchange Web Services request.</span></span>  <br/> |
   
## <a name="remarks"></a><span data-ttu-id="48f7b-138">備考</span><span class="sxs-lookup"><span data-stu-id="48f7b-138">Remarks</span></span>

<span data-ttu-id="48f7b-139">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="48f7b-139">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="48f7b-140">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="48f7b-140">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="48f7b-141">要素情報</span><span class="sxs-lookup"><span data-stu-id="48f7b-141">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="48f7b-142">名前空間</span><span class="sxs-lookup"><span data-stu-id="48f7b-142">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="48f7b-143">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="48f7b-143">Schema Name</span></span>  <br/> |<span data-ttu-id="48f7b-144">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="48f7b-144">Message schema</span></span>  <br/> |
|<span data-ttu-id="48f7b-145">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="48f7b-145">Validation File</span></span>  <br/> |<span data-ttu-id="48f7b-146">messages.xsd</span><span class="sxs-lookup"><span data-stu-id="48f7b-146">messages.xsd</span></span>  <br/> |
|<span data-ttu-id="48f7b-147">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="48f7b-147">Can Be Empty</span></span>  <br/> ||
   
## <a name="see-also"></a><span data-ttu-id="48f7b-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="48f7b-148">See also</span></span>



- [<span data-ttu-id="48f7b-149">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="48f7b-149">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
