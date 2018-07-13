---
title: GetMailTips 操作
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- GetMailTips
api_type:
- schema
ms.assetid: 025483ec-a9f3-4735-8a95-d26e30ea7974
description: GetMailTips 操作は、指定したメールボックスのメール ヒントの情報を取得します。
ms.openlocfilehash: 15c21bef90fdc4cbc6cd65512cdc078fcdf31e60
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19760786"
---
# <a name="getmailtips-operation"></a><span data-ttu-id="f85a5-103">GetMailTips 操作</span><span class="sxs-lookup"><span data-stu-id="f85a5-103">GetMailTips operation</span></span>

<span data-ttu-id="f85a5-104">**GetMailTips**操作は、指定したメールボックスのメール ヒントの情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="f85a5-104">The **GetMailTips** operation gets the mail tips information for the specified mailbox.</span></span> 
  
## <a name="getmailtips-request-example"></a><span data-ttu-id="f85a5-105">GetMailTips 要求の例</span><span class="sxs-lookup"><span data-stu-id="f85a5-105">GetMailTips request example</span></span>

### <a name="description"></a><span data-ttu-id="f85a5-106">説明</span><span class="sxs-lookup"><span data-stu-id="f85a5-106">Description</span></span>

<span data-ttu-id="f85a5-107">クライアントは、要求の XML を構築し、サーバーに送信します。</span><span class="sxs-lookup"><span data-stu-id="f85a5-107">The client constructs the request XML and sends it to the server.</span></span> <span data-ttu-id="f85a5-108">要求のクライアントがメールボックスと、メールのヒントを取得するために送信し、どのようなメール ヒントは、要求されたユーザーを指定します。</span><span class="sxs-lookup"><span data-stu-id="f85a5-108">The request identifies who the client is sending as, the mailbox to retrieve the mail tips for, and what mail tips are requested.</span></span> <span data-ttu-id="f85a5-109">この例では、クライアントは、選択したメールボックスのすべてのメール ヒントが返されることを要求します。</span><span class="sxs-lookup"><span data-stu-id="f85a5-109">In this example, the client requests that all mail tips be returned for the selected mailbox.</span></span>
  
### <a name="code"></a><span data-ttu-id="f85a5-110">コード</span><span class="sxs-lookup"><span data-stu-id="f85a5-110">Code</span></span>

```XML
<?xml version="1.0" encoding="utf-8" ?> 
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
        xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
        xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> 
  <soap:Header> 
    <t:RequestServerVersion Version="Exchange2010" /> 
  </soap:Header> 
  <soap:Body> 
    <GetMailTips xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"> 
      <SendingAs> 
        <t:EmailAddress> user1@contoso.com </t:EmailAddress> 
        <t:RoutingType>SMTP</t:RoutingType> 
      </SendingAs> 
      <Recipients> 
        <t:Mailbox> 
          <t:EmailAddress> user2@contoso.com </t:EmailAddress> 
          <t:RoutingType>SMTP</t:RoutingType> 
        </t:Mailbox> 
      </Recipients> 
      <MailTipsRequested>All</MailTipsRequested> 
    </GetMailTips> 
  </soap:Body> 
</soap:Envelope>
```

### <a name="request-elements"></a><span data-ttu-id="f85a5-111">要素を要求します。</span><span class="sxs-lookup"><span data-stu-id="f85a5-111">Request elements</span></span>

<span data-ttu-id="f85a5-112">要求では、次の要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f85a5-112">The following elements are included in the request:</span></span>
  
- [<span data-ttu-id="f85a5-113">GetMailTips</span><span class="sxs-lookup"><span data-stu-id="f85a5-113">GetMailTips</span></span>](getmailtips.md)
    
- [<span data-ttu-id="f85a5-114">SendingAs</span><span class="sxs-lookup"><span data-stu-id="f85a5-114">SendingAs</span></span>](sendingas.md)
    
- [<span data-ttu-id="f85a5-115">受信者 (ArrayOfRecipientsType)</span><span class="sxs-lookup"><span data-stu-id="f85a5-115">Recipients (ArrayOfRecipientsType)</span></span>](recipients-arrayofrecipientstype.md)
    
- [<span data-ttu-id="f85a5-116">MailTipsRequested</span><span class="sxs-lookup"><span data-stu-id="f85a5-116">MailTipsRequested</span></span>](mailtipsrequested.md)
    
## <a name="successful-getmailtips-response-example"></a><span data-ttu-id="f85a5-117">成功した GetMailTips の応答の例</span><span class="sxs-lookup"><span data-stu-id="f85a5-117">Successful GetMailTips response example</span></span>

### <a name="description"></a><span data-ttu-id="f85a5-118">説明</span><span class="sxs-lookup"><span data-stu-id="f85a5-118">Description</span></span>

<span data-ttu-id="f85a5-119">Simple Object Access Protocol (SOAP) 本文の次の使用例は、 **GetMailTips**要求に正常な応答を示しています。</span><span class="sxs-lookup"><span data-stu-id="f85a5-119">The following Simple Object Access Protocol (SOAP) body example shows a successful response to the **GetMailTips** request.</span></span> 
  
### <a name="code"></a><span data-ttu-id="f85a5-120">コード</span><span class="sxs-lookup"><span data-stu-id="f85a5-120">Code</span></span>

```XML
<?xml version="1.0" encoding="utf-8"?> 
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"> 
  <s:Header> 
    <h:ServerVersionInfo MajorVersion="14" MinorVersion="0" MajorBuildNumber="536" MinorBuildNumber="0" Version="Exchange2010" 
xmlns:h="http://schemas.microsoft.com/exchange/services/2006/types" 
xmlns="http://schemas.microsoft.com/exchange/services/2006/types" 
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
xmlns:xsd="http://www.w3.org/2001/XMLSchema"/> 
  </s:Header> 
  <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"> 
    <GetMailTipsResponse ResponseClass="Success" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"> 
      <ResponseCode>NoError</ResponseCode> 
      <ResponseMessages> 
        <MailTipsResponseMessageType ResponseClass="Success"> 
        <ResponseCode>NoError</ResponseCode> 
        <m:MailTips xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"> 20 / 29 [MS-OXWMT] — v20100517 Mail Tips Web Service Extensions Copyright © 2010 Microsoft Corporation. Release: Monday, May 17, 2010 
          <t:RecipientAddress xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> 
          <t:Name/> 
          <t:EmailAddress>user2@contoso.com</t:EmailAddress> 
          <t:RoutingType>SMTP</t:RoutingType> 
          </t:RecipientAddress> 
          <t:PendingMailTips xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"/> 
          <t:OutOfOffice xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> 
            <t:ReplyBody> 
              <t:Message/> 
            </t:ReplyBody> 
          </t:OutOfOffice> 
          <t:MailboxFull xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">false</t:MailboxFull> 
          <t:CustomMailTip xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">Hello World Mailtips</t:CustomMailTip> 
          <t:TotalMemberCount xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">1</t:TotalMemberCount> 
          <t:ExternalMemberCount xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">0</t:ExternalMemberCount> 
          <t:MaxMessageSize xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">10485760</t:MaxMessageSize> 
          <t:DeliveryRestricted xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">false</t:DeliveryRestricted> 
          <t:IsModerated xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">false</t:IsModerated> 
          <t:InvalidRecipient xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">false</t:InvalidRecipient> 
        </m:MailTips> 
        </MailTipsResponseMessageType> 
      </ResponseMessages> 
    </GetMailTipsResponse> 
  </s:Body> 
</s:Envelope>
```

### <a name="response-elements"></a><span data-ttu-id="f85a5-121">応答の要素</span><span class="sxs-lookup"><span data-stu-id="f85a5-121">Response elements</span></span>

<span data-ttu-id="f85a5-122">応答では、次の要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f85a5-122">The following elements are included in the response:</span></span>
  
- [<span data-ttu-id="f85a5-123">ResponseCode</span><span class="sxs-lookup"><span data-stu-id="f85a5-123">ResponseCode</span></span>](responsecode.md)
    
- [<span data-ttu-id="f85a5-124">メール ヒント</span><span class="sxs-lookup"><span data-stu-id="f85a5-124">MailTips</span></span>](mailtips.md)
    
## <a name="see-also"></a><span data-ttu-id="f85a5-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="f85a5-125">See also</span></span>



[<span data-ttu-id="f85a5-126">Exchange での EWS の操作</span><span class="sxs-lookup"><span data-stu-id="f85a5-126">EWS operations in Exchange</span></span>](ews-operations-in-exchange.md)
  
- [<span data-ttu-id="f85a5-127">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="f85a5-127">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)
