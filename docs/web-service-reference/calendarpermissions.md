---
title: CalendarPermissions
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- CalendarPermissions
api_type:
- schema
ms.assetid: 9b659d83-45fb-42a2-b052-5bc4dbe3854d
description: CalendarPermissions 要素には、フォルダーのアクセス許可を予定表の配列が含まれています。 この要素は、Microsoft Exchange Server 2007 Service Pack 1 (SP1) で導入されました。
ms.openlocfilehash: 3dee635e4449cbb3717f5d2fab8838f3e43102a4
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19759595"
---
# <a name="calendarpermissions"></a><span data-ttu-id="07299-104">CalendarPermissions</span><span class="sxs-lookup"><span data-stu-id="07299-104">CalendarPermissions</span></span>

<span data-ttu-id="07299-105">**CalendarPermissions**要素には、フォルダーのアクセス許可を予定表の配列が含まれています。</span><span class="sxs-lookup"><span data-stu-id="07299-105">The **CalendarPermissions** element contains an array of calendar permissions for a folder.</span></span> <span data-ttu-id="07299-106">この要素は、Microsoft Exchange Server 2007 Service Pack 1 (SP1) で導入されました。</span><span class="sxs-lookup"><span data-stu-id="07299-106">This element was introduced in Microsoft Exchange Server 2007 Service Pack 1 (SP1).</span></span> 
  
```xml
<CalendarPermissions>
   <PermissionSet/>
</CalendarPermissions>
```

 <span data-ttu-id="07299-107">**ArrayOfCalendarPermissionsType**</span><span class="sxs-lookup"><span data-stu-id="07299-107">**ArrayOfCalendarPermissionsType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="07299-108">属性および要素</span><span class="sxs-lookup"><span data-stu-id="07299-108">Attributes and elements</span></span>

<span data-ttu-id="07299-109">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="07299-109">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="07299-110">属性</span><span class="sxs-lookup"><span data-stu-id="07299-110">Attributes</span></span>

<span data-ttu-id="07299-111">なし。</span><span class="sxs-lookup"><span data-stu-id="07299-111">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="07299-112">子要素</span><span class="sxs-lookup"><span data-stu-id="07299-112">Child elements</span></span>

|<span data-ttu-id="07299-113">**要素**</span><span class="sxs-lookup"><span data-stu-id="07299-113">**Element**</span></span>|<span data-ttu-id="07299-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="07299-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="07299-115">CalendarPermission</span><span class="sxs-lookup"><span data-stu-id="07299-115">CalendarPermission</span></span>](calendarpermission.md) <br/> |<span data-ttu-id="07299-116">予定表フォルダーを代理人のユーザーが持つアクセス権を定義します。</span><span class="sxs-lookup"><span data-stu-id="07299-116">Defines the access that a delegate user has to a calendar folder.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="07299-117">親要素</span><span class="sxs-lookup"><span data-stu-id="07299-117">Parent elements</span></span>

|<span data-ttu-id="07299-118">**要素**</span><span class="sxs-lookup"><span data-stu-id="07299-118">**Element**</span></span>|<span data-ttu-id="07299-119">**説明**</span><span class="sxs-lookup"><span data-stu-id="07299-119">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="07299-120">PermissionSet (CalendarPermissionSetType)</span><span class="sxs-lookup"><span data-stu-id="07299-120">PermissionSet (CalendarPermissionSetType)</span></span>](permissionset-calendarpermissionsettype.md) <br/> |<span data-ttu-id="07299-121">予定表フォルダーに対して構成されているすべてのアクセス許可が含まれています。</span><span class="sxs-lookup"><span data-stu-id="07299-121">Contains all the configured permissions for a calendar folder.</span></span> <span data-ttu-id="07299-122">この要素は、Exchange 2007 SP1 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="07299-122">This element was introduced in Exchange 2007 SP1.</span></span>  <br/> |
   
## <a name="remarks"></a><span data-ttu-id="07299-123">備考</span><span class="sxs-lookup"><span data-stu-id="07299-123">Remarks</span></span>

<span data-ttu-id="07299-124">この要素を記述するスキーマは、クライアント アクセス サーバーの役割がインストールされている Microsoft Exchange Server 2007 を実行しているコンピューターの EWS 仮想ディレクトリにあります。</span><span class="sxs-lookup"><span data-stu-id="07299-124">The schema that describes this element is located in the EWS virtual directory of the computer that is running Microsoft Exchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="07299-125">要素情報</span><span class="sxs-lookup"><span data-stu-id="07299-125">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="07299-126">名前空間</span><span class="sxs-lookup"><span data-stu-id="07299-126">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="07299-127">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="07299-127">Schema Name</span></span>  <br/> |<span data-ttu-id="07299-128">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="07299-128">Types schema</span></span>  <br/> |
|<span data-ttu-id="07299-129">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="07299-129">Validation File</span></span>  <br/> |<span data-ttu-id="07299-130">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="07299-130">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="07299-131">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="07299-131">Can be Empty</span></span>  <br/> |<span data-ttu-id="07299-132">False</span><span class="sxs-lookup"><span data-stu-id="07299-132">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="07299-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="07299-133">See also</span></span>



- [<span data-ttu-id="07299-134">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="07299-134">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)


[<span data-ttu-id="07299-135">フォルダー レベルのアクセス許可の設定</span><span class="sxs-lookup"><span data-stu-id="07299-135">Setting Folder-Level Permissions</span></span>](http://msdn.microsoft.com/library/c7530e86-5112-401c-b10a-9c054ae59f07%28Office.15%29.aspx)

