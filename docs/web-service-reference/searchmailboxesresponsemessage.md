---
title: SearchMailboxesResponseMessage
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 6f1bcbfb-d7f6-4fa0-b6f8-681a0b067007
description: SearchMailboxesResponseMessage 要素は、SearchMailboxes 要求の応答メッセージを指定します。
ms.openlocfilehash: b9e16e5b45271b366dd5cc49537085dbc8234da0
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19833297"
---
# <a name="searchmailboxesresponsemessage"></a><span data-ttu-id="660fe-103">SearchMailboxesResponseMessage</span><span class="sxs-lookup"><span data-stu-id="660fe-103">SearchMailboxesResponseMessage</span></span>

<span data-ttu-id="660fe-104">**SearchMailboxesResponseMessage**要素は、 **SearchMailboxes**要求の応答メッセージを指定します。</span><span class="sxs-lookup"><span data-stu-id="660fe-104">The **SearchMailboxesResponseMessage** element specifies the response message for a **SearchMailboxes** request.</span></span> 
  
```XML
<SearchMailboxesResponseMessage>
   <MessageText/>
   <ResponseCode/>
   <DescriptiveLinkKey/>
   <MessageXml/>
   <SearchMailboxesResult/>
</SearchMailboxesResponseMessage>
```

 <span data-ttu-id="660fe-105">**SearchMailboxesResponseMessageType**</span><span class="sxs-lookup"><span data-stu-id="660fe-105">**SearchMailboxesResponseMessageType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="660fe-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="660fe-106">Attributes and elements</span></span>

<span data-ttu-id="660fe-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="660fe-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="660fe-108">属性</span><span class="sxs-lookup"><span data-stu-id="660fe-108">Attributes</span></span>

<span data-ttu-id="660fe-109">なし。</span><span class="sxs-lookup"><span data-stu-id="660fe-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="660fe-110">子要素</span><span class="sxs-lookup"><span data-stu-id="660fe-110">Child elements</span></span>

<span data-ttu-id="660fe-111">[メッセージ テキスト](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md) | [MessageXml](messagexml.md) | [SearchMailboxesResult](searchmailboxesresult.md)</span><span class="sxs-lookup"><span data-stu-id="660fe-111">[MessageText](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md) | [MessageXml](messagexml.md) | [SearchMailboxesResult](searchmailboxesresult.md)</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="660fe-112">親要素</span><span class="sxs-lookup"><span data-stu-id="660fe-112">Parent elements</span></span>

<span data-ttu-id="660fe-113">なし。</span><span class="sxs-lookup"><span data-stu-id="660fe-113">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="660fe-114">備考</span><span class="sxs-lookup"><span data-stu-id="660fe-114">Remarks</span></span>

<span data-ttu-id="660fe-115">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="660fe-115">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="660fe-116">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="660fe-116">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="660fe-117">要素情報</span><span class="sxs-lookup"><span data-stu-id="660fe-117">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="660fe-118">名前空間</span><span class="sxs-lookup"><span data-stu-id="660fe-118">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="660fe-119">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="660fe-119">Schema name</span></span>  <br/> |<span data-ttu-id="660fe-120">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="660fe-120">Messages schema</span></span>  <br/> |
|<span data-ttu-id="660fe-121">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="660fe-121">Validation file</span></span>  <br/> |<span data-ttu-id="660fe-122">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="660fe-122">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="660fe-123">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="660fe-123">Can be empty</span></span>  <br/> ||
   
