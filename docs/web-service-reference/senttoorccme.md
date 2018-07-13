---
title: SentToOrCcMe
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- SentToOrCcMe
api_type:
- schema
ms.assetid: ca43e05d-df37-485b-9276-34678025f2b7
description: SentToOrCcMe 要素は、メールボックスの所有者が ToRecipients または CcRecipients の条件または例外を適用する順序で受信メッセージのプロパティであるかどうかを示します。
ms.openlocfilehash: baed8f71349c9ec06173d0b494ece688f6fc2c5a
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19833361"
---
# <a name="senttoorccme"></a><span data-ttu-id="3386e-103">SentToOrCcMe</span><span class="sxs-lookup"><span data-stu-id="3386e-103">SentToOrCcMe</span></span>

<span data-ttu-id="3386e-104">**SentToOrCcMe**要素は、メールボックスの所有者が**ToRecipients**または**CcRecipients**のいずれかの条件または例外を適用する順序で受信メッセージのプロパティであるかどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="3386e-104">The **SentToOrCcMe** element indicates whether the owner of the mailbox has to be in either a **ToRecipients** or **CcRecipients** property of incoming messages in order for the condition or exception to apply.</span></span> 
  
```XML
<SentToOrCcMe>true | false</SentToOrCcMe>
```

 <span data-ttu-id="3386e-105">**ブール型 (Boolean)**</span><span class="sxs-lookup"><span data-stu-id="3386e-105">**Boolean**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="3386e-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="3386e-106">Attributes and elements</span></span>

<span data-ttu-id="3386e-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="3386e-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="3386e-108">属性</span><span class="sxs-lookup"><span data-stu-id="3386e-108">Attributes</span></span>

<span data-ttu-id="3386e-109">なし。</span><span class="sxs-lookup"><span data-stu-id="3386e-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="3386e-110">子要素</span><span class="sxs-lookup"><span data-stu-id="3386e-110">Child elements</span></span>

<span data-ttu-id="3386e-111">なし。</span><span class="sxs-lookup"><span data-stu-id="3386e-111">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="3386e-112">親要素</span><span class="sxs-lookup"><span data-stu-id="3386e-112">Parent elements</span></span>

|<span data-ttu-id="3386e-113">**要素**</span><span class="sxs-lookup"><span data-stu-id="3386e-113">**Element**</span></span>|<span data-ttu-id="3386e-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="3386e-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="3386e-115">条件</span><span class="sxs-lookup"><span data-stu-id="3386e-115">Conditions</span></span>](conditions.md) <br/> |<span data-ttu-id="3386e-116">条件を表しますが、満たされるとときに、ルールのルールの処理をトリガーします。</span><span class="sxs-lookup"><span data-stu-id="3386e-116">Represents the conditions that, when fulfilled, will trigger the rule actions for a rule.</span></span>  <br/> |
|[<span data-ttu-id="3386e-117">Exceptions</span><span class="sxs-lookup"><span data-stu-id="3386e-117">Exceptions</span></span>](exceptions.md) <br/> |<span data-ttu-id="3386e-118">受信トレイ ルールの使用可能なルールの例外条件をすべてを表します。</span><span class="sxs-lookup"><span data-stu-id="3386e-118">Represents all the available rule exception conditions for an Inbox rule.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="3386e-119">テキスト値</span><span class="sxs-lookup"><span data-stu-id="3386e-119">Text value</span></span>

<span data-ttu-id="3386e-120">テキスト値が**true**のでは、メールボックスの所有者を適用する場合の条件または例外の順序で受信メッセージの**ToRecipients**または**CcRecipients**プロパティにする必要があることを示します。</span><span class="sxs-lookup"><span data-stu-id="3386e-120">A text value of **true** indicates that the owner of the mailbox must be in the **ToRecipients** or **CcRecipients** property of the incoming messages in order for the condition or exception to apply.</span></span> <span data-ttu-id="3386e-121">**False**の値を示します**ToRecipients**または**CcRecipients**プロパティを適用する場合の条件または例外の順序で受信メッセージのでは、メールボックスの所有者でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="3386e-121">A value of **false** indicates that the owner of the mailbox must not be in the **ToRecipients** or **CcRecipients** property of the incoming messages in order for the condition or exception to apply.</span></span> 
  
## <a name="remarks"></a><span data-ttu-id="3386e-122">備考</span><span class="sxs-lookup"><span data-stu-id="3386e-122">Remarks</span></span>

<span data-ttu-id="3386e-123">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="3386e-123">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="3386e-124">要素情報</span><span class="sxs-lookup"><span data-stu-id="3386e-124">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="3386e-125">名前空間</span><span class="sxs-lookup"><span data-stu-id="3386e-125">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="3386e-126">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="3386e-126">Schema Name</span></span>  <br/> |<span data-ttu-id="3386e-127">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="3386e-127">Messages schema</span></span>  <br/> |
|<span data-ttu-id="3386e-128">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="3386e-128">Validation File</span></span>  <br/> |<span data-ttu-id="3386e-129">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="3386e-129">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="3386e-130">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="3386e-130">Can be Empty</span></span>  <br/> |<span data-ttu-id="3386e-131">True</span><span class="sxs-lookup"><span data-stu-id="3386e-131">True</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="3386e-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="3386e-132">See also</span></span>



- [<span data-ttu-id="3386e-133">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="3386e-133">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
