---
title: IsEqualTo
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- IsEqualTo
api_type:
- schema
ms.assetid: 48e7e067-049c-4184-8026-071e6f558e8a
description: IsEqualTo 要素は、プロパティを定数値または別のプロパティを比較し、両者が等しい場合に true に評価する検索式を表します。
ms.openlocfilehash: a7a7deed79c271be74bb2ff16dd86605d468721b
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832013"
---
# <a name="isequalto"></a><span data-ttu-id="1705b-103">IsEqualTo</span><span class="sxs-lookup"><span data-stu-id="1705b-103">IsEqualTo</span></span>

<span data-ttu-id="1705b-104">**IsEqualTo**要素は、プロパティを定数値または別のプロパティを比較し、両者が等しい場合に true に評価する検索式を表します。</span><span class="sxs-lookup"><span data-stu-id="1705b-104">The **IsEqualTo** element represents a search expression that compares a property with either a constant value or another property and evaluates to true if they are equal.</span></span> 
  
```xml
<IsEqualTo>
   <FieldURI/>
   <FieldURIOrConstant/>
</IsEqualTo>
```

 <span data-ttu-id="1705b-105">**IsEqualToType**</span><span class="sxs-lookup"><span data-stu-id="1705b-105">**IsEqualToType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="1705b-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="1705b-106">Attributes and elements</span></span>

<span data-ttu-id="1705b-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="1705b-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="1705b-108">属性</span><span class="sxs-lookup"><span data-stu-id="1705b-108">Attributes</span></span>

<span data-ttu-id="1705b-109">なし。</span><span class="sxs-lookup"><span data-stu-id="1705b-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="1705b-110">子要素</span><span class="sxs-lookup"><span data-stu-id="1705b-110">Child elements</span></span>

|<span data-ttu-id="1705b-111">**要素**</span><span class="sxs-lookup"><span data-stu-id="1705b-111">**Element**</span></span>|<span data-ttu-id="1705b-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="1705b-112">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="1705b-113">FieldURI</span><span class="sxs-lookup"><span data-stu-id="1705b-113">FieldURI</span></span>](fielduri.md) <br/> |<span data-ttu-id="1705b-114">URI によって頻繁に参照されるプロパティを識別します。</span><span class="sxs-lookup"><span data-stu-id="1705b-114">Identifies frequently referenced properties by URI.</span></span>  <br/> |
|[<span data-ttu-id="1705b-115">IndexedFieldURI</span><span class="sxs-lookup"><span data-stu-id="1705b-115">IndexedFieldURI</span></span>](indexedfielduri.md) <br/> |<span data-ttu-id="1705b-116">辞書の個々 のメンバーを識別します。</span><span class="sxs-lookup"><span data-stu-id="1705b-116">Identifies individual members of a dictionary.</span></span>  <br/> |
|[<span data-ttu-id="1705b-117">ExtendedFieldURI</span><span class="sxs-lookup"><span data-stu-id="1705b-117">ExtendedFieldURI</span></span>](extendedfielduri.md) <br/> |<span data-ttu-id="1705b-118">MAPI プロパティを識別します。</span><span class="sxs-lookup"><span data-stu-id="1705b-118">Identifies MAPI properties.</span></span>  <br/> |
|[<span data-ttu-id="1705b-119">FieldURIOrConstant</span><span class="sxs-lookup"><span data-stu-id="1705b-119">FieldURIOrConstant</span></span>](fielduriorconstant.md) <br/> |<span data-ttu-id="1705b-120">プロパティまたは別のプロパティを比較するときに使用する定数値のいずれかを表します。</span><span class="sxs-lookup"><span data-stu-id="1705b-120">Represents either a property or a constant value to be used when comparing with another property.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="1705b-121">親要素</span><span class="sxs-lookup"><span data-stu-id="1705b-121">Parent elements</span></span>

|<span data-ttu-id="1705b-122">**要素**</span><span class="sxs-lookup"><span data-stu-id="1705b-122">**Element**</span></span>|<span data-ttu-id="1705b-123">**説明**</span><span class="sxs-lookup"><span data-stu-id="1705b-123">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="1705b-124">Restriction</span><span class="sxs-lookup"><span data-stu-id="1705b-124">Restriction</span></span>](restriction.md) <br/> |<span data-ttu-id="1705b-125">制限またはアイテムまたはフォルダーの FindItem/FindFolder、検索フォルダーの操作にフィルターを適用するために使用するクエリを表します。</span><span class="sxs-lookup"><span data-stu-id="1705b-125">Represents the restriction or query that is used to filter items or folders in FindItem/FindFolder and search folder operations.</span></span>  <br/> |
|[<span data-ttu-id="1705b-126">Not</span><span class="sxs-lookup"><span data-stu-id="1705b-126">Not</span></span>](not.md) <br/> |<span data-ttu-id="1705b-127">含まれている検索式のブール値を否定する検索式を表します。</span><span class="sxs-lookup"><span data-stu-id="1705b-127">Represents a search expression that negates the Boolean value of the search expression that it contains.</span></span>  <br/> |
|[<span data-ttu-id="1705b-128">And</span><span class="sxs-lookup"><span data-stu-id="1705b-128">And</span></span>](and.md) <br/> |<span data-ttu-id="1705b-129">使用すると、2 つまたは複数の検索式間で論理 And 演算を実行する検索式を表します。</span><span class="sxs-lookup"><span data-stu-id="1705b-129">Represents a search expression that allows you to perform a Boolean And operation between two or more search expressions.</span></span> <span data-ttu-id="1705b-130">**すべての And に含まれている検索式に**該当**する場合は、And 演算の結果は**</span><span class="sxs-lookup"><span data-stu-id="1705b-130">The result of the And operation is **true** if all of the search expressions contained within the And are **true**.</span></span>  <br/> |
|[<span data-ttu-id="1705b-131">Or</span><span class="sxs-lookup"><span data-stu-id="1705b-131">Or</span></span>](or.md) <br/> |<span data-ttu-id="1705b-132">含まれている検索式に対して論理 OR を実行する検索式を表します。</span><span class="sxs-lookup"><span data-stu-id="1705b-132">Represents a search expression that performs a logical OR on the search expression that it contains.</span></span> <span data-ttu-id="1705b-133">[または](or.md)その子のいずれかの場合は true を返す場合は true。 戻ります。</span><span class="sxs-lookup"><span data-stu-id="1705b-133">[Or](or.md) will return true if any of its children return true.</span></span> <span data-ttu-id="1705b-134">**または**2 つ以上の子が必要です。</span><span class="sxs-lookup"><span data-stu-id="1705b-134">**Or** must have two or more children.</span></span>  <br/> |
   
## <a name="remarks"></a><span data-ttu-id="1705b-135">備考</span><span class="sxs-lookup"><span data-stu-id="1705b-135">Remarks</span></span>

<span data-ttu-id="1705b-136">文字列比較を実行するには、検討、[含まれる](contains.md)要素を使用して一致するパラメーターは、大文字と小文字や空白文字などのオプションが用意されています。</span><span class="sxs-lookup"><span data-stu-id="1705b-136">To perform string comparisons, consider using the [Contains](contains.md) element, as it provides options for matching parameters, such as case and white space.</span></span> <span data-ttu-id="1705b-137">結果の符号を反転するのに[含まれる](contains.md)要素と共に、要素では[なく](not.md)を使用します。</span><span class="sxs-lookup"><span data-stu-id="1705b-137">Use the [Not](not.md) element in conjunction with the [Contains](contains.md) element to negate the result.</span></span> 
  
<span data-ttu-id="1705b-138">MicrosoftExchange Server 2007 がインストールされているクライアント アクセス サーバーの役割を実行しているコンピューターの EWS 仮想ディレクトリには、この要素を記述するスキーマがあります。</span><span class="sxs-lookup"><span data-stu-id="1705b-138">The schema that describes this element is located in the EWS virtual directory of the computer that is running MicrosoftExchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="1705b-139">要素情報</span><span class="sxs-lookup"><span data-stu-id="1705b-139">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="1705b-140">名前空間</span><span class="sxs-lookup"><span data-stu-id="1705b-140">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="1705b-141">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="1705b-141">Schema Name</span></span>  <br/> |<span data-ttu-id="1705b-142">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="1705b-142">Types schema</span></span>  <br/> |
|<span data-ttu-id="1705b-143">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="1705b-143">Validation File</span></span>  <br/> |<span data-ttu-id="1705b-144">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="1705b-144">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="1705b-145">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="1705b-145">Can be Empty</span></span>  <br/> |<span data-ttu-id="1705b-146">False</span><span class="sxs-lookup"><span data-stu-id="1705b-146">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="1705b-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="1705b-147">See also</span></span>



- [<span data-ttu-id="1705b-148">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="1705b-148">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
