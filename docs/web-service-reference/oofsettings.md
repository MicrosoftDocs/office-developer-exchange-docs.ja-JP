---
title: OofSettings
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- OofSettings
api_type:
- schema
ms.assetid: 8f37d174-db11-427c-bbed-fdde754a60c7
description: OofSettings 要素には、Office (OOF) の設定が含まれています。
ms.openlocfilehash: d71f068ff24af22da98b6b4de090ab26d3f74f26
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832649"
---
# <a name="oofsettings"></a><span data-ttu-id="e00bb-103">OofSettings</span><span class="sxs-lookup"><span data-stu-id="e00bb-103">OofSettings</span></span>

<span data-ttu-id="e00bb-104">**OofSettings**要素には、Office (OOF) の設定が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e00bb-104">The **OofSettings** element contains the Out of Office (OOF) settings.</span></span> 
  
[<span data-ttu-id="e00bb-105">GetUserOofSettingsResponse</span><span class="sxs-lookup"><span data-stu-id="e00bb-105">GetUserOofSettingsResponse</span></span>](getuseroofsettingsresponse.md)
  
[<span data-ttu-id="e00bb-106">OofSettings</span><span class="sxs-lookup"><span data-stu-id="e00bb-106">OofSettings</span></span>](oofsettings.md)
  
```xml
<OofSettings>
   <OofState>...</OofState>
   <ExternalAudience>...</ExternalAudience>
   <Duration>...</Duration>
   <InternalReply>...</InternalReply>
   <ExternalReply>...</ExternalReply>
</OofSettings>
```

 <span data-ttu-id="e00bb-107">**UserOofSettings**</span><span class="sxs-lookup"><span data-stu-id="e00bb-107">**UserOofSettings**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="e00bb-108">属性および要素</span><span class="sxs-lookup"><span data-stu-id="e00bb-108">Attributes and elements</span></span>

<span data-ttu-id="e00bb-109">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="e00bb-109">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="e00bb-110">属性</span><span class="sxs-lookup"><span data-stu-id="e00bb-110">Attributes</span></span>

<span data-ttu-id="e00bb-111">なし。</span><span class="sxs-lookup"><span data-stu-id="e00bb-111">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="e00bb-112">子要素</span><span class="sxs-lookup"><span data-stu-id="e00bb-112">Child elements</span></span>

|<span data-ttu-id="e00bb-113">**要素**</span><span class="sxs-lookup"><span data-stu-id="e00bb-113">**Element**</span></span>|<span data-ttu-id="e00bb-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="e00bb-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="e00bb-115">OofState</span><span class="sxs-lookup"><span data-stu-id="e00bb-115">OofState</span></span>](oofstate.md) <br/> |<span data-ttu-id="e00bb-116">ユーザーの不在時の状態が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e00bb-116">Contains the user's OOF state.</span></span>  <br/> |
|[<span data-ttu-id="e00bb-117">ExternalAudience</span><span class="sxs-lookup"><span data-stu-id="e00bb-117">ExternalAudience</span></span>](externalaudience.md) <br/> |<span data-ttu-id="e00bb-118">外部の OOF メッセージを送信するかを決定する値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e00bb-118">Contains a value that determines to whom external OOF messages are sent.</span></span>  <br/> |
|[<span data-ttu-id="e00bb-119">期間 (UserOofSettings)</span><span class="sxs-lookup"><span data-stu-id="e00bb-119">Duration (UserOofSettings)</span></span>](duration-useroofsettings.md) <br/> |<span data-ttu-id="e00bb-120">不在の状態が有効である[OofState](oofstate.md)要素は、**スケジュール**に設定されている場合の期間が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e00bb-120">Contains the duration for which the OOF status is enabled if the [OofState](oofstate.md) element is set to **Scheduled**.</span></span> <span data-ttu-id="e00bb-121">[OofState](oofstate.md)要素は、**有効**または**無効**に設定されている場合、この要素の値は無視されます。</span><span class="sxs-lookup"><span data-stu-id="e00bb-121">If the [OofState](oofstate.md) element is set to **Enabled** or **Disabled**, the value of this element is ignored.</span></span>  <br/> |
|[<span data-ttu-id="e00bb-122">InternalReply</span><span class="sxs-lookup"><span data-stu-id="e00bb-122">InternalReply</span></span>](internalreply.md) <br/> |<span data-ttu-id="e00bb-123">ユーザーのドメインまたは信頼されたドメイン内の他のユーザーに送信される不在時の応答が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e00bb-123">Contains the OOF response sent to other users in the user's domain or trusted domain.</span></span>  <br/> |
|[<span data-ttu-id="e00bb-124">ExternalReply</span><span class="sxs-lookup"><span data-stu-id="e00bb-124">ExternalReply</span></span>](externalreply.md) <br/> |<span data-ttu-id="e00bb-125">受信者のドメインまたは信頼されるドメインの外部アドレスに送信される不在時の応答が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e00bb-125">Contains the OOF response sent to addresses outside the recipient's domain or trusted domains.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="e00bb-126">親要素</span><span class="sxs-lookup"><span data-stu-id="e00bb-126">Parent elements</span></span>

|<span data-ttu-id="e00bb-127">**要素**</span><span class="sxs-lookup"><span data-stu-id="e00bb-127">**Element**</span></span>|<span data-ttu-id="e00bb-128">**説明**</span><span class="sxs-lookup"><span data-stu-id="e00bb-128">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="e00bb-129">GetUserOofSettingsResponse</span><span class="sxs-lookup"><span data-stu-id="e00bb-129">GetUserOofSettingsResponse</span></span>](getuseroofsettingsresponse.md) <br/> |<span data-ttu-id="e00bb-130">応答結果とユーザーの不在時の設定が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e00bb-130">Contains the response results and the OOF settings for a user.</span></span>  <br/> <span data-ttu-id="e00bb-131">この要素への XPath 式は、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="e00bb-131">The following is the XPath expression to this element:</span></span>  <br/>  `/GetUserOofSettingsResponse` <br/> |
   
## <a name="remarks"></a><span data-ttu-id="e00bb-132">備考</span><span class="sxs-lookup"><span data-stu-id="e00bb-132">Remarks</span></span>

<span data-ttu-id="e00bb-133">MicrosoftExchange Server 2007 がインストールされているクライアント アクセス サーバーの役割を実行しているコンピューターの EWS 仮想ディレクトリには、この要素を記述するスキーマがあります。</span><span class="sxs-lookup"><span data-stu-id="e00bb-133">The schema that describes this element is located in the EWS virtual directory of the computer that is running MicrosoftExchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="e00bb-134">要素情報</span><span class="sxs-lookup"><span data-stu-id="e00bb-134">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="e00bb-135">名前空間</span><span class="sxs-lookup"><span data-stu-id="e00bb-135">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="e00bb-136">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="e00bb-136">Schema Name</span></span>  <br/> |<span data-ttu-id="e00bb-137">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="e00bb-137">Messages schema</span></span>  <br/> |
|<span data-ttu-id="e00bb-138">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="e00bb-138">Validation File</span></span>  <br/> |<span data-ttu-id="e00bb-139">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="e00bb-139">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="e00bb-140">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="e00bb-140">Can be Empty</span></span>  <br/> |<span data-ttu-id="e00bb-141">False</span><span class="sxs-lookup"><span data-stu-id="e00bb-141">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="e00bb-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="e00bb-142">See also</span></span>



[<span data-ttu-id="e00bb-143">GetUserOofSettings 操作</span><span class="sxs-lookup"><span data-stu-id="e00bb-143">GetUserOofSettings operation</span></span>](getuseroofsettings-operation.md)
  
[<span data-ttu-id="e00bb-144">SetUserOofSettings 操作</span><span class="sxs-lookup"><span data-stu-id="e00bb-144">SetUserOofSettings operation</span></span>](setuseroofsettings-operation.md)

