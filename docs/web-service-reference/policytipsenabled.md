---
title: PolicyTipsEnabled
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 16409652-21e4-4bd3-9373-67e1882236b4
description: PolicyTipsEnabled 要素は、ポリシーのヒントが有効になっているかどうかを示します。
ms.openlocfilehash: 683131a5cefd6757faf582324f312b01fd9ddb33
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832838"
---
# <a name="policytipsenabled"></a><span data-ttu-id="58a13-103">PolicyTipsEnabled</span><span class="sxs-lookup"><span data-stu-id="58a13-103">PolicyTipsEnabled</span></span>

<span data-ttu-id="58a13-104">**PolicyTipsEnabled**要素は、ポリシーのヒントが有効になっているかどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="58a13-104">The **PolicyTipsEnabled** element indicates whether policy tips are enabled.</span></span> 
  
```XML
<PolicyTipsEnabled> true | false </PolicyTipsEnabled>
```

 <span data-ttu-id="58a13-105">**ブール型 (Boolean)**</span><span class="sxs-lookup"><span data-stu-id="58a13-105">**Boolean**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="58a13-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="58a13-106">Attributes and elements</span></span>

<span data-ttu-id="58a13-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="58a13-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="58a13-108">属性</span><span class="sxs-lookup"><span data-stu-id="58a13-108">Attributes</span></span>

<span data-ttu-id="58a13-109">なし。</span><span class="sxs-lookup"><span data-stu-id="58a13-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="58a13-110">子要素</span><span class="sxs-lookup"><span data-stu-id="58a13-110">Child elements</span></span>

<span data-ttu-id="58a13-111">なし。</span><span class="sxs-lookup"><span data-stu-id="58a13-111">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="58a13-112">親要素</span><span class="sxs-lookup"><span data-stu-id="58a13-112">Parent elements</span></span>

[<span data-ttu-id="58a13-113">MailTipsConfiguration (MailTipsServiceConfiguration)</span><span class="sxs-lookup"><span data-stu-id="58a13-113">MailTipsConfiguration (MailTipsServiceConfiguration)</span></span>](mailtipsconfiguration-mailtipsserviceconfiguration.md)
  
## <a name="text-value"></a><span data-ttu-id="58a13-114">テキスト値</span><span class="sxs-lookup"><span data-stu-id="58a13-114">Text value</span></span>

<span data-ttu-id="58a13-115">の**場合は true** 、 **PolicyTipsEnabled**要素のテキスト値は、メールボックスのポリシーのヒントが有効であることを示します。</span><span class="sxs-lookup"><span data-stu-id="58a13-115">A text value of **true** for the **PolicyTipsEnabled** element indicates that policy tips are enabled for a mailbox.</span></span> <span data-ttu-id="58a13-116">**False**の値は、メールボックスに対して、ポリシー ヒントが無効になっていることを示します。</span><span class="sxs-lookup"><span data-stu-id="58a13-116">A value of **false** indicates that policy tips are not enabled for a mailbox.</span></span> 
  
## <a name="remarks"></a><span data-ttu-id="58a13-117">備考</span><span class="sxs-lookup"><span data-stu-id="58a13-117">Remarks</span></span>

<span data-ttu-id="58a13-118">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="58a13-118">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="58a13-119">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="58a13-119">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="58a13-120">要素情報</span><span class="sxs-lookup"><span data-stu-id="58a13-120">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="58a13-121">名前空間</span><span class="sxs-lookup"><span data-stu-id="58a13-121">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="58a13-122">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="58a13-122">Schema name</span></span>  <br/> |<span data-ttu-id="58a13-123">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="58a13-123">Types schema</span></span>  <br/> |
|<span data-ttu-id="58a13-124">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="58a13-124">Validation file</span></span>  <br/> |<span data-ttu-id="58a13-125">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="58a13-125">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="58a13-126">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="58a13-126">Can be empty</span></span>  <br/> ||
   
