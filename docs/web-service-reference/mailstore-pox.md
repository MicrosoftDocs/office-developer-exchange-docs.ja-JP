---
title: MailStore (POX)
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
ms.assetid: af338f99-9e62-4124-9bff-8d7cc2008161
description: MailStore 要素には、MAPI または HTTP プロトコルを使用してクライアントをユーザーのメールボックスに接続するための仕様が含まれています。
ms.openlocfilehash: 4c82c7b61752cf7d91287a3968f6c642f4943855
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832301"
---
# <a name="mailstore-pox"></a><span data-ttu-id="3ab78-103">MailStore (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-103">MailStore (POX)</span></span>

<span data-ttu-id="3ab78-104">**MailStore**要素には、MAPI または HTTP プロトコルを使用してクライアントをユーザーのメールボックスに接続するための仕様が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3ab78-104">The **MailStore** element contains the specifications for connecting a client to the user's mailbox by using the MAPI/HTTP protocol.</span></span> 
  
[<span data-ttu-id="3ab78-105">(POX) を自動検出</span><span class="sxs-lookup"><span data-stu-id="3ab78-105">AutoDiscover (POX)</span></span>](autodiscover-pox.md)
  
[<span data-ttu-id="3ab78-106">応答 (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-106">Response (POX)</span></span>](response-pox.md)
  
[<span data-ttu-id="3ab78-107">アカウント (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-107">Account (POX)</span></span>](account-pox.md)
  
[<span data-ttu-id="3ab78-108">プロトコル (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-108">Protocol (POX)</span></span>](protocol-pox.md)
  
[<span data-ttu-id="3ab78-109">MailStore (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-109">MailStore (POX)</span></span>](mailstore-pox.md)
  
```XML
<MailStore>
   <ExternalUrl/>
   <InternalUrl/>
</MailStore>
```

## <a name="attributes-and-elements"></a><span data-ttu-id="3ab78-110">属性および要素</span><span class="sxs-lookup"><span data-stu-id="3ab78-110">Attributes and elements</span></span>

<span data-ttu-id="3ab78-111">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="3ab78-111">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="3ab78-112">属性</span><span class="sxs-lookup"><span data-stu-id="3ab78-112">Attributes</span></span>

<span data-ttu-id="3ab78-113">なし。</span><span class="sxs-lookup"><span data-stu-id="3ab78-113">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="3ab78-114">子要素</span><span class="sxs-lookup"><span data-stu-id="3ab78-114">Child elements</span></span>

|<span data-ttu-id="3ab78-115">**要素**</span><span class="sxs-lookup"><span data-stu-id="3ab78-115">**Element**</span></span>|<span data-ttu-id="3ab78-116">**説明**</span><span class="sxs-lookup"><span data-stu-id="3ab78-116">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="3ab78-117">ExternalUrl (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-117">ExternalUrl (POX)</span></span>](externalurl-pox.md) <br/> |<span data-ttu-id="3ab78-118">MAPI または HTTP プロトコルを使用して組織のネットワークの外部からのユーザーのメールボックスにアクセスするために使用する URL が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3ab78-118">Contains the URL that should be used to access the user's mailbox from outside the organization's network by means of the MAPI/HTTP protocol.</span></span>  <br/> |
|[<span data-ttu-id="3ab78-119">InternalUrl (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-119">InternalUrl (POX)</span></span>](internalurl-pox.md) <br/> |<span data-ttu-id="3ab78-120">MAPI または HTTP プロトコルを使用して組織のネットワーク内からのユーザーのメールボックスへのアクセスに使用する URL が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3ab78-120">Contains the URL that should be used to access the user's mailbox from inside the organization's network by means of the MAPI/HTTP protocol.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="3ab78-121">親要素</span><span class="sxs-lookup"><span data-stu-id="3ab78-121">Parent elements</span></span>

|<span data-ttu-id="3ab78-122">**要素**</span><span class="sxs-lookup"><span data-stu-id="3ab78-122">**Element**</span></span>|<span data-ttu-id="3ab78-123">**説明**</span><span class="sxs-lookup"><span data-stu-id="3ab78-123">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="3ab78-124">プロトコル (POX)</span><span class="sxs-lookup"><span data-stu-id="3ab78-124">Protocol (POX)</span></span>](protocol-pox.md) <br/> |<span data-ttu-id="3ab78-125">クライアント アクセス サーバーにクライアントを接続するための仕様が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3ab78-125">Contains the specifications for connecting a client to the Client Access server.</span></span>  <br/> |
   
## <a name="remarks"></a><span data-ttu-id="3ab78-126">備考</span><span class="sxs-lookup"><span data-stu-id="3ab78-126">Remarks</span></span>

<span data-ttu-id="3ab78-127">**MailStore**要素は、"mapiHttp"の**型**の属性値を持つ[プロトコル (POX)](protocol-pox.md)要素を持つ応答に存在します。</span><span class="sxs-lookup"><span data-stu-id="3ab78-127">The **MailStore** element is present in a response that has a [Protocol (POX)](protocol-pox.md) element with a **Type** attribute value of "mapiHttp".</span></span> 
  
<span data-ttu-id="3ab78-128">**MailStore**要素は、ターゲット Exchange Online では、Office 365 の一部として Exchange Online MAPI または HTTP プロトコルを実装するクライアントに使用し、設置型のバージョンの Exchange が始まる 15.00.0847.032 (Exchange Server 2013 SP1) を構築します。</span><span class="sxs-lookup"><span data-stu-id="3ab78-128">The **MailStore** element is available to clients that implement the MAPI/HTTP protocol and target Exchange Online, Exchange Online as part of Office 365, and on-premises versions of Exchange starting with build 15.00.0847.032 (Exchange Server 2013 SP1).</span></span> 
  
## <a name="see-also"></a><span data-ttu-id="3ab78-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="3ab78-129">See also</span></span>



[<span data-ttu-id="3ab78-130">交換の POX の自動検出の XML 要素</span><span class="sxs-lookup"><span data-stu-id="3ab78-130">POX Autodiscover XML elements for Exchange</span></span>](pox-autodiscover-xml-elements-for-exchange.md)
