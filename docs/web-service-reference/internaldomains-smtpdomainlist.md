---
title: InternalDomains (SmtpDomainList)
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- InternalDomains
api_type:
- schema
ms.assetid: 0f2cbb05-338d-4302-8871-a06e78b33f98
description: InternalDomains 要素は、組織の内部の SMTP ドメインのリストを識別します。
ms.openlocfilehash: f37a31f4348a7eb0024656489f249dec349bc67b
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19831953"
---
# <a name="internaldomains-smtpdomainlist"></a><span data-ttu-id="2521a-103">InternalDomains (SmtpDomainList)</span><span class="sxs-lookup"><span data-stu-id="2521a-103">InternalDomains (SmtpDomainList)</span></span>

<span data-ttu-id="2521a-104">**InternalDomains**要素は、組織の内部の SMTP ドメインのリストを識別します。</span><span class="sxs-lookup"><span data-stu-id="2521a-104">The **InternalDomains** element identifies the list of internal SMTP domains of the organization.</span></span> 
  
```XML
<InternalDomains>
   <Domain/>
</InternalDomains>
```

 <span data-ttu-id="2521a-105">**SmtpDomainList**</span><span class="sxs-lookup"><span data-stu-id="2521a-105">**SmtpDomainList**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="2521a-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="2521a-106">Attributes and elements</span></span>

<span data-ttu-id="2521a-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="2521a-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="2521a-108">属性</span><span class="sxs-lookup"><span data-stu-id="2521a-108">Attributes</span></span>

<span data-ttu-id="2521a-109">なし。</span><span class="sxs-lookup"><span data-stu-id="2521a-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="2521a-110">子要素</span><span class="sxs-lookup"><span data-stu-id="2521a-110">Child elements</span></span>

|<span data-ttu-id="2521a-111">**要素**</span><span class="sxs-lookup"><span data-stu-id="2521a-111">**Element**</span></span>|<span data-ttu-id="2521a-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="2521a-112">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="2521a-113">ドメイン</span><span class="sxs-lookup"><span data-stu-id="2521a-113">Domain</span></span>](domain.md) <br/> |<span data-ttu-id="2521a-114">1 つの SMTP ドメインを識別します。</span><span class="sxs-lookup"><span data-stu-id="2521a-114">Identifies a single SMTP domain.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="2521a-115">親要素</span><span class="sxs-lookup"><span data-stu-id="2521a-115">Parent elements</span></span>

|<span data-ttu-id="2521a-116">**要素**</span><span class="sxs-lookup"><span data-stu-id="2521a-116">**Element**</span></span>|<span data-ttu-id="2521a-117">**説明**</span><span class="sxs-lookup"><span data-stu-id="2521a-117">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="2521a-118">MailTipsConfiguration (MailTipsServiceConfiguration)</span><span class="sxs-lookup"><span data-stu-id="2521a-118">MailTipsConfiguration (MailTipsServiceConfiguration)</span></span>](mailtipsconfiguration-mailtipsserviceconfiguration.md) <br/> |<span data-ttu-id="2521a-119">メール ヒント サービスのサービスの構成情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2521a-119">Contains service configuration information for the mail tips service.</span></span>  <br/> |
|[<span data-ttu-id="2521a-120">ProtectionRulesConfiguration</span><span class="sxs-lookup"><span data-stu-id="2521a-120">ProtectionRulesConfiguration</span></span>](protectionrulesconfiguration.md) <br/> |<span data-ttu-id="2521a-121">保護ルールのサービスのサービスの構成情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2521a-121">Contains service configuration information for the protection rules service.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="2521a-122">テキスト値</span><span class="sxs-lookup"><span data-stu-id="2521a-122">Text value</span></span>

<span data-ttu-id="2521a-123">なし。</span><span class="sxs-lookup"><span data-stu-id="2521a-123">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="2521a-124">備考</span><span class="sxs-lookup"><span data-stu-id="2521a-124">Remarks</span></span>

<span data-ttu-id="2521a-125">この要素は必須です。</span><span class="sxs-lookup"><span data-stu-id="2521a-125">This element is required.</span></span> 
  
<span data-ttu-id="2521a-126">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="2521a-126">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="2521a-127">要素情報</span><span class="sxs-lookup"><span data-stu-id="2521a-127">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="2521a-128">名前空間</span><span class="sxs-lookup"><span data-stu-id="2521a-128">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="2521a-129">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="2521a-129">Schema Name</span></span>  <br/> |<span data-ttu-id="2521a-130">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="2521a-130">Types schema</span></span>  <br/> |
|<span data-ttu-id="2521a-131">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="2521a-131">Validation File</span></span>  <br/> |<span data-ttu-id="2521a-132">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="2521a-132">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="2521a-133">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="2521a-133">Can be Empty</span></span>  <br/> |<span data-ttu-id="2521a-134">False</span><span class="sxs-lookup"><span data-stu-id="2521a-134">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="2521a-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="2521a-135">See also</span></span>



- [<span data-ttu-id="2521a-136">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="2521a-136">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
