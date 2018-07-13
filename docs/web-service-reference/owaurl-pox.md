---
title: OWAUrl (POX)
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
localization_priority: Normal
api_type:
- schema
ms.assetid: 450b86a1-1722-49f5-b541-16c1edc3db7a
description: OWAUrl 要素は、URL を説明し、Microsoft Exchange Server 2007 を実行する特定のコンピューターへのアクセスに使用する認証スキーマには、クライアント アクセス サーバーの役割がインストールされている Outlook Web Access をホストしています。
ms.openlocfilehash: 93d03506e2a74aa1b4ef6d2a49ccbda01cfc1f9a
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832673"
---
# <a name="owaurl-pox"></a><span data-ttu-id="7cb35-103">OWAUrl (POX)</span><span class="sxs-lookup"><span data-stu-id="7cb35-103">OWAUrl (POX)</span></span>

<span data-ttu-id="7cb35-104">**OWAUrl**要素は、URL を説明し、Microsoft Exchange Server 2007 を実行する特定のコンピューターへのアクセスに使用する認証スキーマには、クライアント アクセス サーバーの役割がインストールされている Outlook Web Access をホストしています。</span><span class="sxs-lookup"><span data-stu-id="7cb35-104">The **OWAUrl** element describes the URL and authentication schema that is used to access a particular computer that is running Microsoft Exchange Server 2007 that has the Client Access server role installed that hosts Outlook Web Access.</span></span> 
  
[<span data-ttu-id="7cb35-105">(POX) を自動検出</span><span class="sxs-lookup"><span data-stu-id="7cb35-105">AutoDiscover (POX)</span></span>](autodiscover-pox.md)
  
[<span data-ttu-id="7cb35-106">応答 (POX)</span><span class="sxs-lookup"><span data-stu-id="7cb35-106">Response (POX)</span></span>](response-pox.md)
  
[<span data-ttu-id="7cb35-107">アカウント (POX)</span><span class="sxs-lookup"><span data-stu-id="7cb35-107">Account (POX)</span></span>](account-pox.md)
  
[<span data-ttu-id="7cb35-108">プロトコル (POX)</span><span class="sxs-lookup"><span data-stu-id="7cb35-108">Protocol (POX)</span></span>](protocol-pox.md)
  
[<span data-ttu-id="7cb35-109">内部 (POX)</span><span class="sxs-lookup"><span data-stu-id="7cb35-109">Internal (POX)</span></span>](internal-pox.md)
  
[<span data-ttu-id="7cb35-110">OWAUrl (POX)</span><span class="sxs-lookup"><span data-stu-id="7cb35-110">OWAUrl (POX)</span></span>](owaurl-pox.md)
  
```xml
<OWAUrl AuthenticationMethod=""/>
```

## <a name="attributes-and-elements"></a><span data-ttu-id="7cb35-111">属性および要素</span><span class="sxs-lookup"><span data-stu-id="7cb35-111">Attributes and elements</span></span>

<span data-ttu-id="7cb35-112">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="7cb35-112">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="7cb35-113">属性</span><span class="sxs-lookup"><span data-stu-id="7cb35-113">Attributes</span></span>

|<span data-ttu-id="7cb35-114">**属性**</span><span class="sxs-lookup"><span data-stu-id="7cb35-114">**Attribute**</span></span>|<span data-ttu-id="7cb35-115">**説明**</span><span class="sxs-lookup"><span data-stu-id="7cb35-115">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7cb35-116">**AuthenticationMethod**</span><span class="sxs-lookup"><span data-stu-id="7cb35-116">**AuthenticationMethod**</span></span> <br/> |<span data-ttu-id="7cb35-117">Outlook Web Access にアクセスするための認証方法をについて説明します。</span><span class="sxs-lookup"><span data-stu-id="7cb35-117">Describes the authentication methods for accessing Outlook Web Access.</span></span>  <br/> |
   
#### <a name="authenticationmethod-attribute"></a><span data-ttu-id="7cb35-118">AuthenticationMethod 属性</span><span class="sxs-lookup"><span data-stu-id="7cb35-118">AuthenticationMethod Attribute</span></span>

|<span data-ttu-id="7cb35-119">**値**</span><span class="sxs-lookup"><span data-stu-id="7cb35-119">**Value**</span></span>|<span data-ttu-id="7cb35-120">**説明**</span><span class="sxs-lookup"><span data-stu-id="7cb35-120">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7cb35-121">WindowsIntegrated</span><span class="sxs-lookup"><span data-stu-id="7cb35-121">WindowsIntegrated</span></span>  <br/> |<span data-ttu-id="7cb35-122">統合 Windows 認証をします。</span><span class="sxs-lookup"><span data-stu-id="7cb35-122">Integrated Windows authentication.</span></span>  <br/> |
|<span data-ttu-id="7cb35-123">FBA</span><span class="sxs-lookup"><span data-stu-id="7cb35-123">FBA</span></span>  <br/> |<span data-ttu-id="7cb35-124">フォーム ベースの認証です。</span><span class="sxs-lookup"><span data-stu-id="7cb35-124">Forms-based authentication.</span></span>  <br/> |
|<span data-ttu-id="7cb35-125">NTLM</span><span class="sxs-lookup"><span data-stu-id="7cb35-125">NTLM</span></span>  <br/> |<span data-ttu-id="7cb35-126">NTLM 認証します。</span><span class="sxs-lookup"><span data-stu-id="7cb35-126">NTLM authentication.</span></span>  <br/> |
|<span data-ttu-id="7cb35-127">Digest</span><span class="sxs-lookup"><span data-stu-id="7cb35-127">Digest</span></span>  <br/> |<span data-ttu-id="7cb35-128">ダイジェスト認証です。</span><span class="sxs-lookup"><span data-stu-id="7cb35-128">Digest authentication.</span></span>  <br/> |
|<span data-ttu-id="7cb35-129">基本</span><span class="sxs-lookup"><span data-stu-id="7cb35-129">Basic</span></span>  <br/> |<span data-ttu-id="7cb35-130">基本認証。</span><span class="sxs-lookup"><span data-stu-id="7cb35-130">Basic authentication.</span></span>  <br/> |
   
### <a name="child-elements"></a><span data-ttu-id="7cb35-131">子要素</span><span class="sxs-lookup"><span data-stu-id="7cb35-131">Child elements</span></span>

<span data-ttu-id="7cb35-132">なし。</span><span class="sxs-lookup"><span data-stu-id="7cb35-132">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="7cb35-133">親要素</span><span class="sxs-lookup"><span data-stu-id="7cb35-133">Parent elements</span></span>

|<span data-ttu-id="7cb35-134">**要素**</span><span class="sxs-lookup"><span data-stu-id="7cb35-134">**Element**</span></span>|<span data-ttu-id="7cb35-135">**説明**</span><span class="sxs-lookup"><span data-stu-id="7cb35-135">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="7cb35-136">内部 (POX)</span><span class="sxs-lookup"><span data-stu-id="7cb35-136">Internal (POX)</span></span>](internal-pox.md) <br/> |<span data-ttu-id="7cb35-137">ファイアウォールの内側になる場合にクライアントが接続できる Outlook Web アクセス Url のコレクションが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7cb35-137">Contains the collection of Outlook Web Access URLs that a client can connect to when it is inside the firewall.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="7cb35-138">テキスト値</span><span class="sxs-lookup"><span data-stu-id="7cb35-138">Text value</span></span>

<span data-ttu-id="7cb35-139">テキスト値は、クライアント アクセス サーバー上の Outlook Web Access サービスの URL を表します。</span><span class="sxs-lookup"><span data-stu-id="7cb35-139">The text value represents the URL for the Outlook Web Access service on a Client Access server.</span></span>
  
## <a name="see-also"></a><span data-ttu-id="7cb35-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="7cb35-140">See also</span></span>



[<span data-ttu-id="7cb35-141">交換の POX の自動検出の XML 要素</span><span class="sxs-lookup"><span data-stu-id="7cb35-141">POX Autodiscover XML elements for Exchange</span></span>](pox-autodiscover-xml-elements-for-exchange.md)
