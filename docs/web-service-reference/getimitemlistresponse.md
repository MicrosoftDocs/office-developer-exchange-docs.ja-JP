---
title: GetImItemListResponse
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 00422885-0e7b-4b50-a9ca-01f24ff9858f
description: GetImItemListResponse 要素は、GetImItemList 要求への応答を定義します。
ms.openlocfilehash: 60d0d9a36b89ad861a6c17702cf48da901138472
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19760751"
---
# <a name="getimitemlistresponse"></a><span data-ttu-id="5b726-103">GetImItemListResponse</span><span class="sxs-lookup"><span data-stu-id="5b726-103">GetImItemListResponse</span></span>

<span data-ttu-id="5b726-104">**GetImItemListResponse**要素は、 **GetImItemList**要求への応答を定義します。</span><span class="sxs-lookup"><span data-stu-id="5b726-104">The **GetImItemListResponse** element defines a response to a **GetImItemList** request.</span></span> 
  
```XML
<GetImItemListResponse>
   <MessageText/>
   <ResponseCode/>
   <DescriptiveLinkKey/>
   <MessageXml/>
   <ImItemList/>
</GetImItemListResponse>
```

 <span data-ttu-id="5b726-105">**GetImItemListResponseMessageType**</span><span class="sxs-lookup"><span data-stu-id="5b726-105">**GetImItemListResponseMessageType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="5b726-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="5b726-106">Attributes and elements</span></span>

<span data-ttu-id="5b726-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="5b726-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="5b726-108">属性</span><span class="sxs-lookup"><span data-stu-id="5b726-108">Attributes</span></span>

<span data-ttu-id="5b726-109">なし。</span><span class="sxs-lookup"><span data-stu-id="5b726-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="5b726-110">子要素</span><span class="sxs-lookup"><span data-stu-id="5b726-110">Child elements</span></span>

<span data-ttu-id="5b726-111">[メッセージ テキスト](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md) | [MessageXml](messagexml.md) | [ImItemList](imitemlist.md)</span><span class="sxs-lookup"><span data-stu-id="5b726-111">[MessageText](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md) | [MessageXml](messagexml.md) | [ImItemList](imitemlist.md)</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="5b726-112">親要素</span><span class="sxs-lookup"><span data-stu-id="5b726-112">Parent elements</span></span>

<span data-ttu-id="5b726-113">なし。</span><span class="sxs-lookup"><span data-stu-id="5b726-113">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="5b726-114">備考</span><span class="sxs-lookup"><span data-stu-id="5b726-114">Remarks</span></span>

<span data-ttu-id="5b726-115">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="5b726-115">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="5b726-116">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="5b726-116">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="5b726-117">要素情報</span><span class="sxs-lookup"><span data-stu-id="5b726-117">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="5b726-118">名前空間</span><span class="sxs-lookup"><span data-stu-id="5b726-118">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="5b726-119">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="5b726-119">Schema name</span></span>  <br/> |<span data-ttu-id="5b726-120">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="5b726-120">Messages schema</span></span>  <br/> |
|<span data-ttu-id="5b726-121">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="5b726-121">Validation file</span></span>  <br/> |<span data-ttu-id="5b726-122">messages.xsd</span><span class="sxs-lookup"><span data-stu-id="5b726-122">messages.xsd</span></span>  <br/> |
|<span data-ttu-id="5b726-123">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="5b726-123">Can be empty</span></span>  <br/> ||
   
