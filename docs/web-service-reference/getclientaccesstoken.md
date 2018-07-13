---
title: GetClientAccessToken
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 3704435b-9c84-4cff-a2a0-8288b4ca31ae
description: GetClientAccessToken 要素には、クライアントのアクセス トークンを取得するための要求が含まれています。
ms.openlocfilehash: 5b5420c5f44d770293c77c0c0ae5c927da3aa856
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19760652"
---
# <a name="getclientaccesstoken"></a><span data-ttu-id="c092c-103">GetClientAccessToken</span><span class="sxs-lookup"><span data-stu-id="c092c-103">GetClientAccessToken</span></span>

<span data-ttu-id="c092c-104">**GetClientAccessToken**要素には、クライアントのアクセス トークンを取得するための要求が含まれています。</span><span class="sxs-lookup"><span data-stu-id="c092c-104">The **GetClientAccessToken** element contains a request to get a client access token.</span></span> 
  
```XML
<GetClientAccessToken>
   <TokenRequests/>
</GetClientAccessToken>
```

 <span data-ttu-id="c092c-105">**GetClientAccessTokenType**</span><span class="sxs-lookup"><span data-stu-id="c092c-105">**GetClientAccessTokenType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="c092c-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="c092c-106">Attributes and elements</span></span>

<span data-ttu-id="c092c-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="c092c-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="c092c-108">属性</span><span class="sxs-lookup"><span data-stu-id="c092c-108">Attributes</span></span>

<span data-ttu-id="c092c-109">なし。</span><span class="sxs-lookup"><span data-stu-id="c092c-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="c092c-110">子要素</span><span class="sxs-lookup"><span data-stu-id="c092c-110">Child elements</span></span>

[<span data-ttu-id="c092c-111">TokenRequests</span><span class="sxs-lookup"><span data-stu-id="c092c-111">TokenRequests</span></span>](tokenrequests.md)
  
### <a name="parent-elements"></a><span data-ttu-id="c092c-112">親要素</span><span class="sxs-lookup"><span data-stu-id="c092c-112">Parent elements</span></span>

<span data-ttu-id="c092c-113">なし。</span><span class="sxs-lookup"><span data-stu-id="c092c-113">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="c092c-114">備考</span><span class="sxs-lookup"><span data-stu-id="c092c-114">Remarks</span></span>

<span data-ttu-id="c092c-115">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="c092c-115">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="c092c-116">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="c092c-116">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="c092c-117">要素情報</span><span class="sxs-lookup"><span data-stu-id="c092c-117">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="c092c-118">名前空間</span><span class="sxs-lookup"><span data-stu-id="c092c-118">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="c092c-119">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="c092c-119">Schema name</span></span>  <br/> |<span data-ttu-id="c092c-120">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="c092c-120">Messages schema</span></span>  <br/> |
|<span data-ttu-id="c092c-121">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="c092c-121">Validation file</span></span>  <br/> |<span data-ttu-id="c092c-122">messages.xsd</span><span class="sxs-lookup"><span data-stu-id="c092c-122">messages.xsd</span></span>  <br/> |
|<span data-ttu-id="c092c-123">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="c092c-123">Can be empty</span></span>  <br/> ||
   
