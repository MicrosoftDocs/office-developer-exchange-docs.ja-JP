---
title: SearchQueries
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 67328dab-321b-45ad-929e-cd83e65ad87e
description: SearchQueries 要素には、メールボックス、および検索に関連するクエリの一覧が含まれています。
ms.openlocfilehash: 182f1ba63b4226ea4ff6445ae9f039197dec38a5
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/11/2018
ms.locfileid: "19833316"
---
# <a name="searchqueries"></a><span data-ttu-id="a3471-103">SearchQueries</span><span class="sxs-lookup"><span data-stu-id="a3471-103">SearchQueries</span></span>

<span data-ttu-id="a3471-104">**SearchQueries**要素には、メールボックス、および検索に関連するクエリの一覧が含まれています。</span><span class="sxs-lookup"><span data-stu-id="a3471-104">The **SearchQueries** element contains a list of mailboxes and associated queries for discovery search.</span></span> 
  
```XML
<SearchQueries>
   <MailboxQuery/>
</SearchQueries>
```

 ****
## <a name="attributes-and-elements"></a><span data-ttu-id="a3471-105">属性および要素</span><span class="sxs-lookup"><span data-stu-id="a3471-105">Attributes and elements</span></span>

<span data-ttu-id="a3471-106">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="a3471-106">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="a3471-107">属性</span><span class="sxs-lookup"><span data-stu-id="a3471-107">Attributes</span></span>

<span data-ttu-id="a3471-108">なし。</span><span class="sxs-lookup"><span data-stu-id="a3471-108">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="a3471-109">子要素</span><span class="sxs-lookup"><span data-stu-id="a3471-109">Child elements</span></span>

[<span data-ttu-id="a3471-110">MailboxQuery</span><span class="sxs-lookup"><span data-stu-id="a3471-110">MailboxQuery</span></span>](mailboxquery.md)
  
### <a name="parent-elements"></a><span data-ttu-id="a3471-111">親要素</span><span class="sxs-lookup"><span data-stu-id="a3471-111">Parent elements</span></span>

<span data-ttu-id="a3471-112">[SearchMailboxes](searchmailboxes.md) | [SearchMailboxesResult](searchmailboxesresult.md)</span><span class="sxs-lookup"><span data-stu-id="a3471-112">[SearchMailboxes](searchmailboxes.md) | [SearchMailboxesResult](searchmailboxesresult.md)</span></span>
  
## <a name="remarks"></a><span data-ttu-id="a3471-113">備考</span><span class="sxs-lookup"><span data-stu-id="a3471-113">Remarks</span></span>

<span data-ttu-id="a3471-114">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="a3471-114">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="a3471-115">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="a3471-115">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="a3471-116">要素情報</span><span class="sxs-lookup"><span data-stu-id="a3471-116">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="a3471-117">名前空間</span><span class="sxs-lookup"><span data-stu-id="a3471-117">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="a3471-118">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="a3471-118">Schema name</span></span>  <br/> |<span data-ttu-id="a3471-119">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="a3471-119">Messages schema</span></span>  <br/> |
|<span data-ttu-id="a3471-120">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="a3471-120">Validation file</span></span>  <br/> |<span data-ttu-id="a3471-121">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="a3471-121">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="a3471-122">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="a3471-122">Can be empty</span></span>  <br/> ||
   

