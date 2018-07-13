---
title: MailTipsConfiguration (MailTipsServiceConfiguration)
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- MailTipsConfiguration
api_type:
- schema
ms.assetid: 9a34515e-815b-4c61-b118-d5f66b80238f
description: MailTipsConfiguration 要素には、メール ヒントのサービスのサービスの構成情報が含まれています。
ms.openlocfilehash: ea92af3ebb2d2f720e5823c5317d09d5bcdb3978
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832321"
---
# <a name="mailtipsconfiguration-mailtipsserviceconfiguration"></a><span data-ttu-id="da2c8-103">MailTipsConfiguration (MailTipsServiceConfiguration)</span><span class="sxs-lookup"><span data-stu-id="da2c8-103">MailTipsConfiguration (MailTipsServiceConfiguration)</span></span>

<span data-ttu-id="da2c8-104">**MailTipsConfiguration**要素には、メール ヒントのサービスのサービスの構成情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da2c8-104">The **MailTipsConfiguration** element contains service configuration information for the mail tips service.</span></span> 
  
```XML
<MailTipsConfiguration>
   <MailTipsEnabled/>
   <MaxRecipientsPerGetMailTipsRequest/>
   <MaxMessageSize/>
   <LargeAudienceThreshold/>
   <ShowExternalRecipientCount/>
   <InternalDomains/>
</MailTipsConfiguration>
```

 <span data-ttu-id="da2c8-105">**MailTipsServiceConfiguration**</span><span class="sxs-lookup"><span data-stu-id="da2c8-105">**MailTipsServiceConfiguration**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="da2c8-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="da2c8-106">Attributes and elements</span></span>

<span data-ttu-id="da2c8-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="da2c8-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="da2c8-108">属性</span><span class="sxs-lookup"><span data-stu-id="da2c8-108">Attributes</span></span>

<span data-ttu-id="da2c8-109">なし。</span><span class="sxs-lookup"><span data-stu-id="da2c8-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="da2c8-110">子要素</span><span class="sxs-lookup"><span data-stu-id="da2c8-110">Child elements</span></span>

|<span data-ttu-id="da2c8-111">**要素**</span><span class="sxs-lookup"><span data-stu-id="da2c8-111">**Element**</span></span>|<span data-ttu-id="da2c8-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="da2c8-112">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="da2c8-113">MailTipsEnabled</span><span class="sxs-lookup"><span data-stu-id="da2c8-113">MailTipsEnabled</span></span>](mailtipsenabled.md) <br/> |<span data-ttu-id="da2c8-114">メール ヒントのサービスが利用可能かどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="da2c8-114">Indicates whether the mail tips service is available.</span></span> <span data-ttu-id="da2c8-115">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="da2c8-115">This element is required.</span></span>  <br/> |
|[<span data-ttu-id="da2c8-116">MaxRecipientsPerGetMailTipsRequest</span><span class="sxs-lookup"><span data-stu-id="da2c8-116">MaxRecipientsPerGetMailTipsRequest</span></span>](maxrecipientspergetmailtipsrequest.md) <br/> |<span data-ttu-id="da2c8-117">[GetMailTips 操作](getmailtips-operation.md)に渡すことができる受信者の最大数を示します。</span><span class="sxs-lookup"><span data-stu-id="da2c8-117">Indicates the maximum number of recipients that can be passed to the [GetMailTips operation](getmailtips-operation.md).</span></span> <span data-ttu-id="da2c8-118">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="da2c8-118">This element is required.</span></span>  <br/> |
|[<span data-ttu-id="da2c8-119">MaxMessageSize</span><span class="sxs-lookup"><span data-stu-id="da2c8-119">MaxMessageSize</span></span>](maxmessagesize.md) <br/> |<span data-ttu-id="da2c8-120">受信者が受け入れることができるメッセージの最大サイズを表します。</span><span class="sxs-lookup"><span data-stu-id="da2c8-120">Represents the maximum message size a recipient can accept.</span></span> <span data-ttu-id="da2c8-121">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="da2c8-121">This element is required.</span></span>  <br/> |
|[<span data-ttu-id="da2c8-122">LargeAudienceThreshold</span><span class="sxs-lookup"><span data-stu-id="da2c8-122">LargeAudienceThreshold</span></span>](largeaudiencethreshold.md) <br/> |<span data-ttu-id="da2c8-123">クライアントの多くの人々 のしきい値を表します。</span><span class="sxs-lookup"><span data-stu-id="da2c8-123">Represents the large audience threshold for a client.</span></span> <span data-ttu-id="da2c8-124">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="da2c8-124">This element is required.</span></span>  <br/> |
|[<span data-ttu-id="da2c8-125">ShowExternalRecipientCount</span><span class="sxs-lookup"><span data-stu-id="da2c8-125">ShowExternalRecipientCount</span></span>](showexternalrecipientcount.md) <br/> |<span data-ttu-id="da2c8-126">[GetMailTips 操作](getmailtips-operation.md)のコンシューマーにメッセージの宛先を外部の受信者の数を示すメール ヒントを表示するかどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="da2c8-126">Indicates whether consumers of the [GetMailTips operation](getmailtips-operation.md) have to show mail tips that indicate the number of external recipients to which a message is addressed.</span></span> <span data-ttu-id="da2c8-127">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="da2c8-127">This element is required.</span></span>  <br/> |
|[<span data-ttu-id="da2c8-128">InternalDomains (SmtpDomainList)</span><span class="sxs-lookup"><span data-stu-id="da2c8-128">InternalDomains (SmtpDomainList)</span></span>](internaldomains-smtpdomainlist.md) <br/> |<span data-ttu-id="da2c8-129">組織の内部の SMTP ドメインのリストを識別します。</span><span class="sxs-lookup"><span data-stu-id="da2c8-129">Identifies the list of internal SMTP domains of the organization.</span></span> <span data-ttu-id="da2c8-130">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="da2c8-130">This element is required.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="da2c8-131">親要素</span><span class="sxs-lookup"><span data-stu-id="da2c8-131">Parent elements</span></span>

|<span data-ttu-id="da2c8-132">**要素**</span><span class="sxs-lookup"><span data-stu-id="da2c8-132">**Element**</span></span>|<span data-ttu-id="da2c8-133">**説明**</span><span class="sxs-lookup"><span data-stu-id="da2c8-133">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="da2c8-134">ServiceConfigurationResponseMessageType</span><span class="sxs-lookup"><span data-stu-id="da2c8-134">ServiceConfigurationResponseMessageType</span></span>](serviceconfigurationresponsemessagetype.md) <br/> |<span data-ttu-id="da2c8-135">サービス構成の設定が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da2c8-135">Contains service configuration settings.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="da2c8-136">テキスト値</span><span class="sxs-lookup"><span data-stu-id="da2c8-136">Text value</span></span>

<span data-ttu-id="da2c8-137">なし。</span><span class="sxs-lookup"><span data-stu-id="da2c8-137">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="da2c8-138">備考</span><span class="sxs-lookup"><span data-stu-id="da2c8-138">Remarks</span></span>

<span data-ttu-id="da2c8-139">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="da2c8-139">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="da2c8-140">要素情報</span><span class="sxs-lookup"><span data-stu-id="da2c8-140">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="da2c8-141">名前空間</span><span class="sxs-lookup"><span data-stu-id="da2c8-141">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="da2c8-142">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="da2c8-142">Schema Name</span></span>  <br/> |<span data-ttu-id="da2c8-143">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="da2c8-143">Messages schema</span></span>  <br/> |
|<span data-ttu-id="da2c8-144">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="da2c8-144">Validation File</span></span>  <br/> |<span data-ttu-id="da2c8-145">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="da2c8-145">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="da2c8-146">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="da2c8-146">Can be Empty</span></span>  <br/> |<span data-ttu-id="da2c8-147">False</span><span class="sxs-lookup"><span data-stu-id="da2c8-147">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="da2c8-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="da2c8-148">See also</span></span>



- [<span data-ttu-id="da2c8-149">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="da2c8-149">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
