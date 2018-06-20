---
title: DirectReports
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: ab88739f-9018-4887-ae46-f1471242628c
description: DirectReports 要素には、取引先担当者の直属の部下を識別するための SMTP 情報が含まれています。
ms.openlocfilehash: b82ccff76c506ddfa6a0fc9ef805d579be64f5c6
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/11/2018
ms.locfileid: "19760053"
---
# <a name="directreports"></a><span data-ttu-id="079f5-103">DirectReports</span><span class="sxs-lookup"><span data-stu-id="079f5-103">DirectReports</span></span>

<span data-ttu-id="079f5-104">**DirectReports**要素には、取引先担当者の直属の部下を識別するための SMTP 情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="079f5-104">The **DirectReports** element contains SMTP information that identifies the direct reports of a contact.</span></span> 
  
```XML
<DirectReports/>
```

 <span data-ttu-id="079f5-105">**SingleRecipientType**</span><span class="sxs-lookup"><span data-stu-id="079f5-105">**SingleRecipientType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="079f5-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="079f5-106">Attributes and elements</span></span>

<span data-ttu-id="079f5-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="079f5-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="079f5-108">属性</span><span class="sxs-lookup"><span data-stu-id="079f5-108">Attributes</span></span>

<span data-ttu-id="079f5-109">なし。</span><span class="sxs-lookup"><span data-stu-id="079f5-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="079f5-110">子要素</span><span class="sxs-lookup"><span data-stu-id="079f5-110">Child elements</span></span>

<span data-ttu-id="079f5-111">なし。</span><span class="sxs-lookup"><span data-stu-id="079f5-111">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="079f5-112">親要素</span><span class="sxs-lookup"><span data-stu-id="079f5-112">Parent elements</span></span>

|<span data-ttu-id="079f5-113">**要素名**</span><span class="sxs-lookup"><span data-stu-id="079f5-113">**Element name**</span></span>|<span data-ttu-id="079f5-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="079f5-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="079f5-115">Contact</span><span class="sxs-lookup"><span data-stu-id="079f5-115">Contact</span></span>](contact.md) <br/> |<span data-ttu-id="079f5-116">表すには、Exchange ストア内のアイテムがお問い合わせください。</span><span class="sxs-lookup"><span data-stu-id="079f5-116">Represents contact item in the Exchange store.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="079f5-117">テキスト値</span><span class="sxs-lookup"><span data-stu-id="079f5-117">Text value</span></span>

<span data-ttu-id="079f5-118">なし。</span><span class="sxs-lookup"><span data-stu-id="079f5-118">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="079f5-119">備考</span><span class="sxs-lookup"><span data-stu-id="079f5-119">Remarks</span></span>

<span data-ttu-id="079f5-120">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="079f5-120">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
<span data-ttu-id="079f5-121">この要素は Exchange Server 2010 Service Pack 2 (SP2) で導入されました。</span><span class="sxs-lookup"><span data-stu-id="079f5-121">This element was introduced in Exchange Server 2010 Service Pack 2 (SP2).</span></span>
  
## <a name="element-information"></a><span data-ttu-id="079f5-122">要素情報</span><span class="sxs-lookup"><span data-stu-id="079f5-122">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="079f5-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="079f5-123">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="079f5-124">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="079f5-124">Schema name</span></span>  <br/> |<span data-ttu-id="079f5-125">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="079f5-125">Types schema</span></span>  <br/> |
|<span data-ttu-id="079f5-126">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="079f5-126">Validation file</span></span>  <br/> |<span data-ttu-id="079f5-127">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="079f5-127">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="079f5-128">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="079f5-128">Can be empty</span></span>  <br/> |<span data-ttu-id="079f5-129">False</span><span class="sxs-lookup"><span data-stu-id="079f5-129">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="079f5-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="079f5-130">See also</span></span>

- [<span data-ttu-id="079f5-131">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="079f5-131">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
- [<span data-ttu-id="079f5-132">連絡先 (Exchange Web サービス) を作成します。</span><span class="sxs-lookup"><span data-stu-id="079f5-132">Creating Contacts (Exchange Web Services)</span></span>](http://msdn.microsoft.com/library/4845917e-70d1-481c-bbd7-011ec6571789%28Office.15%29.aspx)

