---
title: NumberOfMembersWithConflict
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- NumberOfMembersWithConflict
api_type:
- schema
ms.assetid: e61154f7-d262-43ec-b2bf-1ba6804b28dc
description: NumberOfMembersWithConflict 要素は、提案された会議の時間との競合を持つ配布リストのメンバーの数を表します。 この要素では、ビジー、不在時、または仮承諾のステータスを持っているメンバーを表します。
ms.openlocfilehash: 227783b4bed32686e8e098f88498fe8ebb25e3cc
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832634"
---
# <a name="numberofmemberswithconflict"></a><span data-ttu-id="3e7f4-104">NumberOfMembersWithConflict</span><span class="sxs-lookup"><span data-stu-id="3e7f4-104">NumberOfMembersWithConflict</span></span>

<span data-ttu-id="3e7f4-105">**NumberOfMembersWithConflict**要素は、提案された会議の時間との競合を持つ配布リストのメンバーの数を表します。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-105">The **NumberOfMembersWithConflict** element represents the number of distribution list members who have a conflict with a suggested meeting time.</span></span> <span data-ttu-id="3e7f4-106">この要素では、**ビジー**、**不在時**、または**仮承諾**のステータスを持っているメンバーを表します。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-106">This element represents members who have a status of **Busy**, **OOF**, or **Tentative**.</span></span>
  
[<span data-ttu-id="3e7f4-107">GetUserAvailabilityResponse</span><span class="sxs-lookup"><span data-stu-id="3e7f4-107">GetUserAvailabilityResponse</span></span>](getuseravailabilityresponse.md)
  
[<span data-ttu-id="3e7f4-108">SuggestionsResponse</span><span class="sxs-lookup"><span data-stu-id="3e7f4-108">SuggestionsResponse</span></span>](suggestionsresponse.md)
  
[<span data-ttu-id="3e7f4-109">SuggestionDayResultArray</span><span class="sxs-lookup"><span data-stu-id="3e7f4-109">SuggestionDayResultArray</span></span>](suggestiondayresultarray.md)
  
[<span data-ttu-id="3e7f4-110">SuggestionDayResult</span><span class="sxs-lookup"><span data-stu-id="3e7f4-110">SuggestionDayResult</span></span>](suggestiondayresult.md)
  
[<span data-ttu-id="3e7f4-111">SuggestionArray</span><span class="sxs-lookup"><span data-stu-id="3e7f4-111">SuggestionArray</span></span>](suggestionarray.md)
  
[<span data-ttu-id="3e7f4-112">提案</span><span class="sxs-lookup"><span data-stu-id="3e7f4-112">Suggestion</span></span>](suggestion.md)
  
[<span data-ttu-id="3e7f4-113">AttendeeConflictDataArray</span><span class="sxs-lookup"><span data-stu-id="3e7f4-113">AttendeeConflictDataArray</span></span>](attendeeconflictdataarray.md)
  
[<span data-ttu-id="3e7f4-114">GroupAttendeeConflictData</span><span class="sxs-lookup"><span data-stu-id="3e7f4-114">GroupAttendeeConflictData</span></span>](groupattendeeconflictdata.md)
  
[<span data-ttu-id="3e7f4-115">NumberOfMembersWithConflict</span><span class="sxs-lookup"><span data-stu-id="3e7f4-115">NumberOfMembersWithConflict</span></span>](numberofmemberswithconflict.md)
  
```xml
<NumberOfMembersWithConflict>...</NumberOfMembersWithConflict>
```

 <span data-ttu-id="3e7f4-116">**int**</span><span class="sxs-lookup"><span data-stu-id="3e7f4-116">**int**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="3e7f4-117">属性および要素</span><span class="sxs-lookup"><span data-stu-id="3e7f4-117">Attributes and elements</span></span>

<span data-ttu-id="3e7f4-118">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-118">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="3e7f4-119">属性</span><span class="sxs-lookup"><span data-stu-id="3e7f4-119">Attributes</span></span>

<span data-ttu-id="3e7f4-120">なし。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-120">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="3e7f4-121">子要素</span><span class="sxs-lookup"><span data-stu-id="3e7f4-121">Child elements</span></span>

<span data-ttu-id="3e7f4-122">なし。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-122">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="3e7f4-123">親要素</span><span class="sxs-lookup"><span data-stu-id="3e7f4-123">Parent elements</span></span>

|<span data-ttu-id="3e7f4-124">**要素**</span><span class="sxs-lookup"><span data-stu-id="3e7f4-124">**Element**</span></span>|<span data-ttu-id="3e7f4-125">**説明**</span><span class="sxs-lookup"><span data-stu-id="3e7f4-125">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="3e7f4-126">GroupAttendeeConflictData</span><span class="sxs-lookup"><span data-stu-id="3e7f4-126">GroupAttendeeConflictData</span></span>](groupattendeeconflictdata.md) <br/> |<span data-ttu-id="3e7f4-127">利用するユーザーの数、競合を持っているユーザーの数および提案された会議の配布リストの利用可能時間情報を持っていないユーザーの数についての集計の競合に関する情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-127">Contains aggregate conflict information about the number of users who are available, the number of users who have conflicts, and the number of users who do not have availability information in a distribution list for a suggested meeting time.</span></span>  <br/> <span data-ttu-id="3e7f4-128">この要素への XPath 式は、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-128">The following is the XPath expression to this element:</span></span>  <br/>  `/GetUserAvailabilityResponse/SuggestionsResponse/SuggestionDayResultArray/SuggestionDayResult[i]/SuggestionArray/Suggestion[i]/AttendeeConflictDataArray/GroupAttendeeConflictData[i]` <br/> |
   
## <a name="remarks"></a><span data-ttu-id="3e7f4-129">備考</span><span class="sxs-lookup"><span data-stu-id="3e7f4-129">Remarks</span></span>

<span data-ttu-id="3e7f4-130">MicrosoftExchange Server 2007 がインストールされているクライアント アクセス サーバーの役割を実行しているコンピューターの EWS 仮想ディレクトリには、この要素を記述するスキーマがあります。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-130">The schema that describes this element is located in the EWS virtual directory of the computer that is running MicrosoftExchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="3e7f4-131">要素情報</span><span class="sxs-lookup"><span data-stu-id="3e7f4-131">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="3e7f4-132">名前空間</span><span class="sxs-lookup"><span data-stu-id="3e7f4-132">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="3e7f4-133">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="3e7f4-133">Schema Name</span></span>  <br/> |<span data-ttu-id="3e7f4-134">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="3e7f4-134">Types schema</span></span>  <br/> |
|<span data-ttu-id="3e7f4-135">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="3e7f4-135">Validation File</span></span>  <br/> |<span data-ttu-id="3e7f4-136">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="3e7f4-136">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="3e7f4-137">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="3e7f4-137">Can be Empty</span></span>  <br/> |<span data-ttu-id="3e7f4-138">False</span><span class="sxs-lookup"><span data-stu-id="3e7f4-138">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="3e7f4-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="3e7f4-139">See also</span></span>



[<span data-ttu-id="3e7f4-140">GetUserAvailability 操作</span><span class="sxs-lookup"><span data-stu-id="3e7f4-140">GetUserAvailability operation</span></span>](getuseravailability-operation.md)
  
[<span data-ttu-id="3e7f4-141">GetUserAvailabilityResponse</span><span class="sxs-lookup"><span data-stu-id="3e7f4-141">GetUserAvailabilityResponse</span></span>](getuseravailabilityresponse.md)


[<span data-ttu-id="3e7f4-142">ユーザーの状態を取得します。</span><span class="sxs-lookup"><span data-stu-id="3e7f4-142">Getting User Availability</span></span>](http://msdn.microsoft.com/library/d4133fcb-9b0f-4e6b-aadf-a389da83516a%28Office.15%29.aspx)
