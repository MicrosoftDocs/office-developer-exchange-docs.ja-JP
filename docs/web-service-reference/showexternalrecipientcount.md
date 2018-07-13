---
title: ShowExternalRecipientCount
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- ShowExternalRecipientCount
api_type:
- schema
ms.assetid: db28dbcb-d051-4e5c-a9c2-4b8d5149b4e1
description: ShowExternalRecipientCount 要素は、メッセージの宛先を外部の受信者の数を示すメール ヒントを表示するのには、GetMailTips 操作の消費者があるかどうかを示します。
ms.openlocfilehash: 1fd3ceb629689c560dc60afe01f0413602f79a0d
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19833491"
---
# <a name="showexternalrecipientcount"></a><span data-ttu-id="62d40-103">ShowExternalRecipientCount</span><span class="sxs-lookup"><span data-stu-id="62d40-103">ShowExternalRecipientCount</span></span>

<span data-ttu-id="62d40-104">**ShowExternalRecipientCount**要素は、メッセージの宛先を外部の受信者の数を示すメール ヒントを表示するのには、 [GetMailTips 操作](getmailtips-operation.md)の消費者があるかどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="62d40-104">The **ShowExternalRecipientCount** element indicates whether consumers of the [GetMailTips operation](getmailtips-operation.md) have to show mail tips that indicate the number of external recipients to which a message is addressed.</span></span> 
  
```XML
<ShowExternalRecipientCount>true | false</ShowExternalRecipientCount>
```

 <span data-ttu-id="62d40-105">**ブール型 (Boolean)**</span><span class="sxs-lookup"><span data-stu-id="62d40-105">**Boolean**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="62d40-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="62d40-106">Attributes and elements</span></span>

<span data-ttu-id="62d40-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="62d40-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="62d40-108">属性</span><span class="sxs-lookup"><span data-stu-id="62d40-108">Attributes</span></span>

<span data-ttu-id="62d40-109">なし。</span><span class="sxs-lookup"><span data-stu-id="62d40-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="62d40-110">子要素</span><span class="sxs-lookup"><span data-stu-id="62d40-110">Child elements</span></span>

<span data-ttu-id="62d40-111">なし。</span><span class="sxs-lookup"><span data-stu-id="62d40-111">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="62d40-112">親要素</span><span class="sxs-lookup"><span data-stu-id="62d40-112">Parent elements</span></span>

|<span data-ttu-id="62d40-113">**要素**</span><span class="sxs-lookup"><span data-stu-id="62d40-113">**Element**</span></span>|<span data-ttu-id="62d40-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="62d40-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="62d40-115">MailTipsConfiguration (MailTipsServiceConfiguration)</span><span class="sxs-lookup"><span data-stu-id="62d40-115">MailTipsConfiguration (MailTipsServiceConfiguration)</span></span>](mailtipsconfiguration-mailtipsserviceconfiguration.md) <br/> |<span data-ttu-id="62d40-116">メール ヒント サービスのサービスの構成情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="62d40-116">Contains service configuration information for the mail tips service.</span></span>  <br/> |
   
## <a name="text-value"></a><span data-ttu-id="62d40-117">テキスト値</span><span class="sxs-lookup"><span data-stu-id="62d40-117">Text value</span></span>

<span data-ttu-id="62d40-118">この要素のテキスト値は、 [GetMailTips 操作](getmailtips-operation.md)のコンシューマーは、メッセージの宛先とする外部の受信者の数を示すメール ヒントを表示する必要がある場合**は true**です。</span><span class="sxs-lookup"><span data-stu-id="62d40-118">The text value of this element is **true** if consumers of the [GetMailTips operation](getmailtips-operation.md) have to show mail tips that indicate the number of external recipients to which a message is addressed.</span></span> <span data-ttu-id="62d40-119">[GetMailTips 操作](getmailtips-operation.md)のコンシューマーは、メッセージの宛先を外部の受信者の数を示すメール ヒントを表示する必要はない場合に、値は**false**です。</span><span class="sxs-lookup"><span data-stu-id="62d40-119">The value is **false** if consumers of the [GetMailTips operation](getmailtips-operation.md) do not have to show mail tips that indicate the number of external recipients to which a message is addressed.</span></span> 
  
## <a name="remarks"></a><span data-ttu-id="62d40-120">備考</span><span class="sxs-lookup"><span data-stu-id="62d40-120">Remarks</span></span>

<span data-ttu-id="62d40-121">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="62d40-121">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="62d40-122">要素情報</span><span class="sxs-lookup"><span data-stu-id="62d40-122">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="62d40-123">名前空間</span><span class="sxs-lookup"><span data-stu-id="62d40-123">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="62d40-124">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="62d40-124">Schema Name</span></span>  <br/> |<span data-ttu-id="62d40-125">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="62d40-125">Types schema</span></span>  <br/> |
|<span data-ttu-id="62d40-126">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="62d40-126">Validation File</span></span>  <br/> |<span data-ttu-id="62d40-127">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="62d40-127">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="62d40-128">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="62d40-128">Can be Empty</span></span>  <br/> |<span data-ttu-id="62d40-129">False</span><span class="sxs-lookup"><span data-stu-id="62d40-129">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="62d40-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="62d40-130">See also</span></span>



[<span data-ttu-id="62d40-131">GetMailTips 操作</span><span class="sxs-lookup"><span data-stu-id="62d40-131">GetMailTips operation</span></span>](getmailtips-operation.md)


- [<span data-ttu-id="62d40-132">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="62d40-132">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
