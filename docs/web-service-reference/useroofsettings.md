---
title: UserOofSettings
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- UserOofSettings
api_type:
- schema
ms.assetid: 0a95ca63-660e-4cc0-82e4-3f74fb4ae21c
description: UserOofSettings 要素は、不在 (OOF) 設定を指定します。
ms.openlocfilehash: 417c3d5061a6229d41eb57f72e89f03213acf460
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44461906"
---
# <a name="useroofsettings"></a><span data-ttu-id="0b75e-103">UserOofSettings</span><span class="sxs-lookup"><span data-stu-id="0b75e-103">UserOofSettings</span></span>

<span data-ttu-id="0b75e-104">**Useroofsettings**要素は、不在 (OOF) 設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-104">The **UserOofSettings** element specifies the Out of Office (OOF) settings.</span></span> 
  
[<span data-ttu-id="0b75e-105">SetUserOofSettingsRequest</span><span class="sxs-lookup"><span data-stu-id="0b75e-105">SetUserOofSettingsRequest</span></span>](setuseroofsettingsrequest.md)
  
[<span data-ttu-id="0b75e-106">UserOofSettings</span><span class="sxs-lookup"><span data-stu-id="0b75e-106">UserOofSettings</span></span>](useroofsettings.md)
  
```xml
<UserOofSettings>
   <OofState>...</OofState>
   <ExternalAudience>...</ExternalAudience>
   <Duration>...</Duration>
   <InternalReply>...</InternalReply>
   <ExternalReply>...</ExternalReply>
</UserOofSettings>
```

 <span data-ttu-id="0b75e-107">**UserOofSettings**</span><span class="sxs-lookup"><span data-stu-id="0b75e-107">**UserOofSettings**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="0b75e-108">属性と要素</span><span class="sxs-lookup"><span data-stu-id="0b75e-108">Attributes and elements</span></span>

<span data-ttu-id="0b75e-109">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-109">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="0b75e-110">属性</span><span class="sxs-lookup"><span data-stu-id="0b75e-110">Attributes</span></span>

<span data-ttu-id="0b75e-111">なし。</span><span class="sxs-lookup"><span data-stu-id="0b75e-111">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="0b75e-112">子要素</span><span class="sxs-lookup"><span data-stu-id="0b75e-112">Child elements</span></span>

|<span data-ttu-id="0b75e-113">**Element**</span><span class="sxs-lookup"><span data-stu-id="0b75e-113">**Element**</span></span>|<span data-ttu-id="0b75e-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="0b75e-114">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="0b75e-115">OofState</span><span class="sxs-lookup"><span data-stu-id="0b75e-115">OofState</span></span>](oofstate.md) <br/> |<span data-ttu-id="0b75e-116">ユーザーの不在時の状態を設定します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-116">Sets the user's OOF state.</span></span>  <br/> |
|[<span data-ttu-id="0b75e-117">ExternalAudience</span><span class="sxs-lookup"><span data-stu-id="0b75e-117">ExternalAudience</span></span>](externalaudience.md) <br/> |<span data-ttu-id="0b75e-118">外部の OOF メッセージを送信するユーザーを指定する値を設定または格納します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-118">Sets or contains a value that determines to whom external OOF messages are sent.</span></span>  <br/> |
|[<span data-ttu-id="0b75e-119">Duration (UserOofSettings)</span><span class="sxs-lookup"><span data-stu-id="0b75e-119">Duration (UserOofSettings)</span></span>](duration-useroofsettings.md) <br/> |<span data-ttu-id="0b75e-120">[Oofstate](oofstate.md)要素が [**スケジュール済み**] に設定されている場合に、不在時の状態を有効にする期間を指定します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-120">Specifies the duration for which the OOF status is enabled if the [OofState](oofstate.md) element is set to **Scheduled**.</span></span> <span data-ttu-id="0b75e-121">[Oofstate](oofstate.md)要素が**Enabled**または**Disabled**に設定されている場合、この要素の値は無視されます。</span><span class="sxs-lookup"><span data-stu-id="0b75e-121">If the [OofState](oofstate.md) element is set to **Enabled** or **Disabled**, the value of this element is ignored.</span></span>  <br/> |
|[<span data-ttu-id="0b75e-122">InternalReply</span><span class="sxs-lookup"><span data-stu-id="0b75e-122">InternalReply</span></span>](internalreply.md) <br/> |<span data-ttu-id="0b75e-123">ユーザーのドメインまたは信頼されたドメイン内の他のユーザーに送信される OOF 応答を格納します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-123">Contains the OOF response sent to other users in the user's domain or trusted domains.</span></span>  <br/> |
|[<span data-ttu-id="0b75e-124">ExternalReply</span><span class="sxs-lookup"><span data-stu-id="0b75e-124">ExternalReply</span></span>](externalreply.md) <br/> |<span data-ttu-id="0b75e-125">受信者のドメインまたは信頼されたドメイン外のアドレスに送信される OOF 応答を含みます。</span><span class="sxs-lookup"><span data-stu-id="0b75e-125">Contains the OOF response sent to addresses outside the recipient's domain or trusted domains.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="0b75e-126">親要素</span><span class="sxs-lookup"><span data-stu-id="0b75e-126">Parent elements</span></span>

|<span data-ttu-id="0b75e-127">**要素**</span><span class="sxs-lookup"><span data-stu-id="0b75e-127">**Element**</span></span>|<span data-ttu-id="0b75e-128">**説明**</span><span class="sxs-lookup"><span data-stu-id="0b75e-128">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="0b75e-129">SetUserOofSettingsRequest</span><span class="sxs-lookup"><span data-stu-id="0b75e-129">SetUserOofSettingsRequest</span></span>](setuseroofsettingsrequest.md) <br/> |<span data-ttu-id="0b75e-130">メールボックスユーザーの不在時の設定とメッセージを設定するために使用する引数が含まれています。</span><span class="sxs-lookup"><span data-stu-id="0b75e-130">Contains the arguments used to set a mailbox user's OOF settings and messages.</span></span>  <br/> <span data-ttu-id="0b75e-131">この要素の XPath 式を次に示します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-131">The following is the XPath expression to this element:</span></span>  <br/>  `/SetUserOofSettingsRequest` <br/> |
   
## <a name="remarks"></a><span data-ttu-id="0b75e-132">注釈</span><span class="sxs-lookup"><span data-stu-id="0b75e-132">Remarks</span></span>

<span data-ttu-id="0b75e-133">この要素を記述するスキーマは、クライアントアクセスサーバーの役割がインストールされている Microsoft Exchange Server 2007 を実行しているコンピューターの EWS 仮想ディレクトリにあります。</span><span class="sxs-lookup"><span data-stu-id="0b75e-133">The schema that describes this element is located in the EWS virtual directory of the computer that is running MicrosoftExchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="example"></a><span data-ttu-id="0b75e-134">例</span><span class="sxs-lookup"><span data-stu-id="0b75e-134">Example</span></span>

<span data-ttu-id="0b75e-135">SetUserOofSettings 要求の次の例では、OoFState を**有効**に設定し、不在時の時間を10日間に設定し、内部および外部の oof メッセージを設定します。</span><span class="sxs-lookup"><span data-stu-id="0b75e-135">The following example of a SetUserOofSettings request sets the OoFState to **Enabled**, sets the duration of OOF for 10 days, and sets the internal and external OOF messages.</span></span>
  
```xml
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <SetUserOofSettingsRequest xmlns="https://schemas.microsoft.com/exchange/services/2006/messages">
      <Mailbox xmlns="https://schemas.microsoft.com/exchange/services/2006/types">
        <Name>David Alexander</Name>
        <Address>someone@example.com</Address>
        <RoutingType>SMTP</RoutingType>
      </Mailbox>
      <UserOofSettings xmlns="https://schemas.microsoft.com/exchange/services/2006/types">
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

## <a name="element-information"></a><span data-ttu-id="0b75e-136">要素の情報</span><span class="sxs-lookup"><span data-stu-id="0b75e-136">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="0b75e-137">Namespace</span><span class="sxs-lookup"><span data-stu-id="0b75e-137">Namespace</span></span>  <br/> |https://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="0b75e-138">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="0b75e-138">Schema Name</span></span>  <br/> |<span data-ttu-id="0b75e-139">メッセージスキーマ</span><span class="sxs-lookup"><span data-stu-id="0b75e-139">messages schema</span></span>  <br/> |
|<span data-ttu-id="0b75e-140">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="0b75e-140">Validation File</span></span>  <br/> |<span data-ttu-id="0b75e-141">メッセージ .xsd</span><span class="sxs-lookup"><span data-stu-id="0b75e-141">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="0b75e-142">空に設定可能</span><span class="sxs-lookup"><span data-stu-id="0b75e-142">Can be Empty</span></span>  <br/> |<span data-ttu-id="0b75e-143">正しくない</span><span class="sxs-lookup"><span data-stu-id="0b75e-143">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="0b75e-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="0b75e-144">See also</span></span>

- [<span data-ttu-id="0b75e-145">SetUserOofSettings 操作</span><span class="sxs-lookup"><span data-stu-id="0b75e-145">SetUserOofSettings operation</span></span>](setuseroofsettings-operation.md)

