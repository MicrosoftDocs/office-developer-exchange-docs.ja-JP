---
title: OofState
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- OofState
api_type:
- schema
ms.assetid: 3c486a38-06da-4382-ad20-664d067d76ac
description: 取得またはユーザーの Office (OOF) の状態を設定する OofState 要素を使用します。
ms.openlocfilehash: f97c050aec102b384fa4d98e6dee43befd4dc9ca
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19832651"
---
# <a name="oofstate"></a><span data-ttu-id="ffadf-103">OofState</span><span class="sxs-lookup"><span data-stu-id="ffadf-103">OofState</span></span>

<span data-ttu-id="ffadf-104">取得またはユーザーの Office (OOF) の状態を設定する**OofState**要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="ffadf-104">The **OofState** element is used to get or set the user's Out of Office (OOF) state.</span></span> 
  
```xml
<OofState>Disabled or Enabled or Scheduled</OofState>
```

 <span data-ttu-id="ffadf-105">**OofState**</span><span class="sxs-lookup"><span data-stu-id="ffadf-105">**OofState**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="ffadf-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="ffadf-106">Attributes and elements</span></span>

<span data-ttu-id="ffadf-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="ffadf-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="ffadf-108">属性</span><span class="sxs-lookup"><span data-stu-id="ffadf-108">Attributes</span></span>

<span data-ttu-id="ffadf-109">なし。</span><span class="sxs-lookup"><span data-stu-id="ffadf-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="ffadf-110">子要素</span><span class="sxs-lookup"><span data-stu-id="ffadf-110">Child elements</span></span>

<span data-ttu-id="ffadf-111">なし。</span><span class="sxs-lookup"><span data-stu-id="ffadf-111">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="ffadf-112">親要素</span><span class="sxs-lookup"><span data-stu-id="ffadf-112">Parent elements</span></span>

|<span data-ttu-id="ffadf-113">**要素**</span><span class="sxs-lookup"><span data-stu-id="ffadf-113">**Element**</span></span>|<span data-ttu-id="ffadf-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="ffadf-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="ffadf-115">UserOofSettings</span><span class="sxs-lookup"><span data-stu-id="ffadf-115">UserOofSettings</span></span>](useroofsettings.md) <br/> |<span data-ttu-id="ffadf-116">不在の設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="ffadf-116">Specifies the OOF settings.</span></span>  <br/> <span data-ttu-id="ffadf-117">この要素への XPath 式は、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="ffadf-117">The following is the XPath expression to this element:</span></span>  <br/>  `/SetUserOofSettingsRequest/UserOofSettings` <br/> |
|[<span data-ttu-id="ffadf-118">OofSettings</span><span class="sxs-lookup"><span data-stu-id="ffadf-118">OofSettings</span></span>](oofsettings.md) <br/> |<span data-ttu-id="ffadf-119">不在の設定が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ffadf-119">Contains the OOF settings.</span></span>  <br/> <span data-ttu-id="ffadf-120">この要素への XPath 式は、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="ffadf-120">The following is the XPath expression to this element:</span></span>  <br/>  `/GetUserOofSettingsResponse/OofSettings` <br/> |
   
## <a name="text-value"></a><span data-ttu-id="ffadf-121">テキスト値</span><span class="sxs-lookup"><span data-stu-id="ffadf-121">Text value</span></span>

<span data-ttu-id="ffadf-122">テキスト値は、 **OofState**要素の必要があります。</span><span class="sxs-lookup"><span data-stu-id="ffadf-122">A text value is required for the **OofState** element.</span></span> <span data-ttu-id="ffadf-123">次の一覧には、この要素の有効な値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ffadf-123">The following list contains the possible values for this element:</span></span> 
  
- <span data-ttu-id="ffadf-124">**無効になっています。**</span><span class="sxs-lookup"><span data-stu-id="ffadf-124">**Disabled**</span></span>
    
- <span data-ttu-id="ffadf-125">**Enabled**</span><span class="sxs-lookup"><span data-stu-id="ffadf-125">**Enabled**</span></span>
    
- <span data-ttu-id="ffadf-126">**スケジュール**</span><span class="sxs-lookup"><span data-stu-id="ffadf-126">**Scheduled**</span></span>
    
<span data-ttu-id="ffadf-127">**[スケジュール済]** の値は、[期間 (UserOofSettings)](duration-useroofsettings.md)の要素で指定された期間内に、不在の状態が**有効**に設定されていることを示します。</span><span class="sxs-lookup"><span data-stu-id="ffadf-127">A value of **Scheduled** indicates that the OOF status is set to **Enabled** during a time period identified by the [Duration (UserOofSettings)](duration-useroofsettings.md) element.</span></span> 
  
## <a name="remarks"></a><span data-ttu-id="ffadf-128">備考</span><span class="sxs-lookup"><span data-stu-id="ffadf-128">Remarks</span></span>

<span data-ttu-id="ffadf-129">SetUsersOofSettingRequest メッセージと、GetUserOofSettingResponse メッセージの両方では、この要素が必要です。</span><span class="sxs-lookup"><span data-stu-id="ffadf-129">This element is required in both the SetUsersOofSettingRequest message and the GetUserOofSettingResponse message.</span></span>
  
<span data-ttu-id="ffadf-130">MicrosoftExchange Server 2007 がインストールされているクライアント アクセス サーバーの役割を実行しているコンピューターの EWS 仮想ディレクトリには、この要素を記述するスキーマがあります。</span><span class="sxs-lookup"><span data-stu-id="ffadf-130">The schema that describes this element is located in the EWS virtual directory of the computer that is running MicrosoftExchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="example"></a><span data-ttu-id="ffadf-131">例</span><span class="sxs-lookup"><span data-stu-id="ffadf-131">Example</span></span>

<span data-ttu-id="ffadf-132">SetUserOofSettings 要求の次の使用例は、 **OofState**を有効にします。</span><span class="sxs-lookup"><span data-stu-id="ffadf-132">The following example of a SetUserOofSettings request enables the **OofState**.</span></span>
  
```
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <SetUserOofSettingsRequest xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <Mailbox xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
        <Name>David Alexander</Name>
        <Address>someone@example.com</Address>
        <RoutingType>SMTP</RoutingType>
      </Mailbox>
      <UserOofSettings xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
        <OofState>Enabled</OofState>
        <ExternalAudience>All</ExternalAudience>
        <Duration>
          <StartTime>2005-10-05T00:00:00</StartTime>
          <EndTime>2005-10-25T00:00:00</EndTime>
        </Duration>
        <InternalReply>
          <Message>I am out of office.  This is my internal reply.</Message>
        </InternalReply>
        <ExternalReply>
          <Message>I am out of office. This is my external reply.</Message>
        </ExternalReply>
      </UserOofSettings>
    </SetUserOofSettingsRequest>
  </soap:Body>
</soap:Envelope>
```

## <a name="element-information"></a><span data-ttu-id="ffadf-133">要素情報</span><span class="sxs-lookup"><span data-stu-id="ffadf-133">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="ffadf-134">名前空間</span><span class="sxs-lookup"><span data-stu-id="ffadf-134">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="ffadf-135">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="ffadf-135">Schema Name</span></span>  <br/> |<span data-ttu-id="ffadf-136">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="ffadf-136">Types schema</span></span>  <br/> |
|<span data-ttu-id="ffadf-137">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="ffadf-137">Validation File</span></span>  <br/> |<span data-ttu-id="ffadf-138">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="ffadf-138">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="ffadf-139">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="ffadf-139">Can be Empty</span></span>  <br/> |<span data-ttu-id="ffadf-140">False</span><span class="sxs-lookup"><span data-stu-id="ffadf-140">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="ffadf-141">関連項目</span><span class="sxs-lookup"><span data-stu-id="ffadf-141">See also</span></span>



[<span data-ttu-id="ffadf-142">GetUserOofSettings 操作</span><span class="sxs-lookup"><span data-stu-id="ffadf-142">GetUserOofSettings operation</span></span>](getuseroofsettings-operation.md)
  
[<span data-ttu-id="ffadf-143">SetUserOofSettings 操作</span><span class="sxs-lookup"><span data-stu-id="ffadf-143">SetUserOofSettings operation</span></span>](setuseroofsettings-operation.md)
