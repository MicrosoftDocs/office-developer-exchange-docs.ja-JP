---
title: ルーム
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- Rooms
api_type:
- schema
ms.assetid: 57b6079a-3d83-4429-861e-c551e9e1a991
description: ルームの要素は、会議室を表す 1 つまたは複数の要素の一覧です。
ms.openlocfilehash: e95112d3d565da29c309e45710ffc0ea9b4d5064
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/11/2018
ms.locfileid: "19833249"
---
# <a name="rooms"></a><span data-ttu-id="5d7b8-103">ルーム</span><span class="sxs-lookup"><span data-stu-id="5d7b8-103">Rooms</span></span>

<span data-ttu-id="5d7b8-104">**ルーム**の要素は、会議室を表す 1 つまたは複数の要素の一覧です。</span><span class="sxs-lookup"><span data-stu-id="5d7b8-104">The **Rooms** element is a list of one or more elements that represent meeting rooms.</span></span> 
  
[<span data-ttu-id="5d7b8-105">GetRoomsResponse</span><span class="sxs-lookup"><span data-stu-id="5d7b8-105">GetRoomsResponse</span></span>](getroomsresponse.md)
  
[<span data-ttu-id="5d7b8-106">ルーム</span><span class="sxs-lookup"><span data-stu-id="5d7b8-106">Rooms</span></span>](rooms.md)
  
```xml
<Rooms>   <Room/></Rooms>
```

 <span data-ttu-id="5d7b8-107">**ArrayOfRoomsType**</span><span class="sxs-lookup"><span data-stu-id="5d7b8-107">**ArrayOfRoomsType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="5d7b8-108">属性および要素</span><span class="sxs-lookup"><span data-stu-id="5d7b8-108">Attributes and elements</span></span>

<span data-ttu-id="5d7b8-109">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="5d7b8-109">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="5d7b8-110">属性</span><span class="sxs-lookup"><span data-stu-id="5d7b8-110">Attributes</span></span>

<span data-ttu-id="5d7b8-111">なし。</span><span class="sxs-lookup"><span data-stu-id="5d7b8-111">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="5d7b8-112">子要素</span><span class="sxs-lookup"><span data-stu-id="5d7b8-112">Child elements</span></span>

|<span data-ttu-id="5d7b8-113">**要素**</span><span class="sxs-lookup"><span data-stu-id="5d7b8-113">**Element**</span></span>|<span data-ttu-id="5d7b8-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="5d7b8-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="5d7b8-115">ルーム</span><span class="sxs-lookup"><span data-stu-id="5d7b8-115">Room</span></span>](room.md) <br/> |<span data-ttu-id="5d7b8-116">電子メール アドレスと、会議室を表す表示名を定義します。</span><span class="sxs-lookup"><span data-stu-id="5d7b8-116">Defines an e-mail address and display name that represents a meeting room.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="5d7b8-117">親要素</span><span class="sxs-lookup"><span data-stu-id="5d7b8-117">Parent elements</span></span>

|<span data-ttu-id="5d7b8-118">**要素**</span><span class="sxs-lookup"><span data-stu-id="5d7b8-118">**Element**</span></span>|<span data-ttu-id="5d7b8-119">**説明**</span><span class="sxs-lookup"><span data-stu-id="5d7b8-119">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="5d7b8-120">GetRoomsResponse</span><span class="sxs-lookup"><span data-stu-id="5d7b8-120">GetRoomsResponse</span></span>](getroomsresponse.md) <br/> ||
   
## <a name="remarks"></a><span data-ttu-id="5d7b8-121">備考</span><span class="sxs-lookup"><span data-stu-id="5d7b8-121">Remarks</span></span>

<span data-ttu-id="5d7b8-122">この要素を記述するスキーマは、クライアント アクセス サーバーの役割がインストールされている Microsoft Exchange Server を実行しているコンピューターの EWS 仮想ディレクトリにあります。</span><span class="sxs-lookup"><span data-stu-id="5d7b8-122">The schema that describes this element is located in the EWS virtual directory of the computer that is running Microsoft Exchange Server that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="5d7b8-123">要素情報</span><span class="sxs-lookup"><span data-stu-id="5d7b8-123">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="5d7b8-124">名前空間</span><span class="sxs-lookup"><span data-stu-id="5d7b8-124">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="5d7b8-125">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="5d7b8-125">Schema Name</span></span>  <br/> |<span data-ttu-id="5d7b8-126">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="5d7b8-126">Messages schema</span></span>  <br/> |
|<span data-ttu-id="5d7b8-127">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="5d7b8-127">Validation File</span></span>  <br/> |<span data-ttu-id="5d7b8-128">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="5d7b8-128">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="5d7b8-129">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="5d7b8-129">Can be Empty</span></span>  <br/> |<span data-ttu-id="5d7b8-130">False</span><span class="sxs-lookup"><span data-stu-id="5d7b8-130">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="5d7b8-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="5d7b8-131">See also</span></span>



[<span data-ttu-id="5d7b8-132">GetRooms 操作</span><span class="sxs-lookup"><span data-stu-id="5d7b8-132">GetRooms operation</span></span>](getrooms-operation.md)


- [<span data-ttu-id="5d7b8-133">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="5d7b8-133">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)

