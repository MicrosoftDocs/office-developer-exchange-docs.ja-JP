---
title: GetHoldOnMailboxes 操作
manager: sethgros
ms.date: 01/24/2020
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 9157f329-80b4-4cd0-a158-378064966ae6
description: GetHoldOnMailboxes EWS 操作についての情報を検索します。
ms.openlocfilehash: 867f38be87e60af8708eeb0b9d0e3ac8eee6ff64
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44457733"
---
# <a name="getholdonmailboxes-operation"></a>GetHoldOnMailboxes 操作

> [!IMPORTANT]
> 2020年4月1日以降、Exchange Online では GetHoldOnMailboxes 操作が使用できなくなります。 この操作は、Exchange Server のオンプレミスバージョンでは影響を受けません。 詳細については、「 [Exchange Online の従来の電子情報開示ツールの廃止](https://docs.microsoft.com/microsoft-365/compliance/legacy-ediscovery-retirement#getsearchablemailboxes-setholdonmailboxes-and-getholdonmailboxes-operations-in-the-ews-api)」を参照してください。

**GetHoldOnMailboxes** EWS 操作についての情報を検索します。 
  
**GetHoldOnMailboxes**操作は、特定の保留中のメールボックスと、それに関連付けられている保留クエリを取得します。 
  
この操作は Exchange Server 2013 で導入されました。
  
## <a name="using-the-getholdonmailboxes-operation"></a>GetHoldOnMailboxes 操作の使用

**GetHoldOnMailboxes**操作は、特定のホールドの下に配置されるメールボックス、各保留リストに関連付けられた保留クエリに関する情報、各メールボックスの保留状態に関する情報をクライアントに提供します。 クエリベースの保持を含む、メールボックス保持の詳細については、TechNet の「[インプレースホールド](https://technet.microsoft.com/library/ff637980%28v=exchg.150%29)」を参照してください。 
  
### <a name="getholdonmailboxes-operation-soap-headers"></a>GetHoldOnMailboxes 操作の SOAP ヘッダー

**GetHoldOnMailboxes**操作では、次の表に示す SOAP ヘッダーを使用できます。 
  
|**ヘッダー名**|**要素**|**説明**|
|:-----|:-----|:-----|
|**Get-managementrole** <br/> |[Get-managementrole](managementrole.md) <br/> |発信者が要求を行うために必要なサーバーの役割を指定します。 このヘッダーは、要求に適用されます。  <br/> |
|**RequestVersion** <br/> |[RequestServerVersion](requestserverversion.md) <br/> |操作要求のスキーマバージョンを識別します。 このヘッダーは、要求に適用されます。  <br/> |
|**ServerVersion** <br/> |[ServerVersionInfo](serverversioninfo.md) <br/> |要求に応答したサーバーのバージョンを識別します。 このヘッダーは応答に適用されます。  <br/> |
   
## <a name="getholdonmailboxes-operation-request-example-get-mailbox-hold-information"></a>GetHoldOnMailboxes 操作要求の例: メールボックスの保持情報を取得します。

次の**GetHoldOnMailboxes**操作要求の例は、HoldId2 メールボックスホールドのメールボックス保持情報を取得する方法を示しています。 
  
```XML
<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="https://schemas.microsoft.com/exchange/services/2006/messages">
   <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" />
   </soap:Header>
   <soap:Body >
      <m:GetHoldOnMailboxes>
         <m:HoldId>HoldId2</m:HoldId>
      </m:GetHoldOnMailboxes>
   </soap:Body>
</soap:Envelope>

```

要求 SOAP 本文には、次の要素が含まれています。
  
- [GetHoldOnMailboxes](getholdonmailboxes.md)
    
- [HoldId](holdid.md)
    
## <a name="successful-getholdonmailboxes-operation-response"></a>成功した GetHoldOnMailboxes 操作の応答

次の例は、HoldId2 メールボックスホールドのメールボックス保持情報を取得するための**GetHoldOnMailboxes**操作要求に対する正常な応答を示しています。 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
   <s:Header>
      <h:ServerVersionInfo MajorVersion="15" 
                           MinorVersion="0" 
                           MajorBuildNumber="526" 
                           MinorBuildNumber="0" 
                           Version="Exchange2013" 
                           xmlns:h="https://schemas.microsoft.com/exchange/services/2006/types" 
                           xmlns="https://schemas.microsoft.com/exchange/services/2006/types" 
                           xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
                           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>
   </s:Header>
   <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:xsd="http://www.w3.org/2001/XMLSchema">
      <GetHoldOnMailboxesResponse ResponseClass="Success" 
                                  xmlns="https://schemas.microsoft.com/exchange/services/2006/messages">
         <ResponseCode>NoError</ResponseCode>
         <MailboxHoldResult>
            <HoldId xmlns="https://schemas.microsoft.com/exchange/services/2006/types">HoldId2</HoldId>
            <Query xmlns="https://schemas.microsoft.com/exchange/services/2006/types">test</Query>
            <MailboxHoldStatuses xmlns="https://schemas.microsoft.com/exchange/services/2006/types">
               <MailboxHoldStatus>
                  <Mailbox>/o=First Organization/ou=Exchange Administrative Group (FYDIBPDLT)/cn=Recipients/cn=ecc0fd98c2cadf-Willi</Mailbox>
                  <Status>OnHold</Status>
                  <AdditionalInfo/>
               </MailboxHoldStatus>
               <MailboxHoldStatus>
                  <Mailbox>/o=First Organization/ou=Exchange Administrative Group (FYDIBPDLT)/cn=Recipients/cn=dasdat341q8de95-Micha</Mailbox>
                  <Status>OnHold</Status>
                  <AdditionalInfo/>
               </MailboxHoldStatus>
            </MailboxHoldStatuses>
         </MailboxHoldResult>
      </GetHoldOnMailboxesResponse>
   </s:Body>
</s:Envelope>

```

応答 SOAP 本文には、次の要素が含まれています。
  
- [GetHoldOnMailboxesResponse](getholdonmailboxesresponse.md)
    
- [ResponseCode](responsecode.md)
    
- [MailboxHoldResult](mailboxholdresult.md)
    
- [HoldId](holdid.md)
    
- [Query](query.md)
    
- [MailboxHoldStatuses](mailboxholdstatuses.md)
    
- [MailboxHoldStatus](mailboxholdstatus.md)
    
- [メールボックス (文字列)](mailbox-string.md)
    
- [状態 (HoldStatusType)](status-holdstatustype.md)
    
- [AdditionalInfo](additionalinfo.md)
    
## <a name="getholdonmailboxes-operation-error-response"></a>GetHoldOnMailboxes 操作エラー応答

次の例は、 **GetHoldOnMailboxes**操作要求へのエラー応答を示しています。 これは、削除された保留リストを取得するための要求に対する応答です。 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
   <s:Header>
      <h:ServerVersionInfo MajorVersion="15" 
                           MinorVersion="0" 
                           MajorBuildNumber="526" 
                           MinorBuildNumber="0" 
                           Version="Exchange2013" 
                           xmlns:h="https://schemas.microsoft.com/exchange/services/2006/types" 
                           xmlns="https://schemas.microsoft.com/exchange/services/2006/types" 
                           xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
                           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>
   </s:Header>
   <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xmlns:xsd="http://www.w3.org/2001/XMLSchema">
      <GetHoldOnMailboxesResponse ResponseClass="Error" 
                                  xmlns="https://schemas.microsoft.com/exchange/services/2006/messages">
         <MessageText>The specific query-based hold is not found.</MessageText>
         <ResponseCode>ErrorMailboxHoldNotFound</ResponseCode>
         <DescriptiveLinkKey>0</DescriptiveLinkKey>
      </GetHoldOnMailboxesResponse>
   </s:Body>
</s:Envelope>

```

エラー応答 SOAP 本文には、次の要素が含まれています。
  
- [GetHoldOnMailboxesResponse](getholdonmailboxesresponse.md)
    
- [MessageText](messagetext.md)
    
- [ResponseCode](responsecode.md)
    
- [DescriptiveLinkKey](descriptivelinkkey.md)
    
EWS で汎用的でこの操作に固有の追加のエラーコードについては、「応答」[を参照してください。](responsecode.md)
  
## <a name="see-also"></a>関連項目

- [Exchange での EWS 操作](ews-operations-in-exchange.md)
    
- [Getsearchablemailemail箱操作](getsearchablemailboxes-operation.md)
    
- [SearchMailboxes ボックスの操作](searchmailboxes-operation.md)
    
- [SetHoldOnMailboxes 操作](setholdonmailboxes-operation.md)
    
- [GetDiscoverySearchConfiguration 操作](getdiscoverysearchconfiguration-operation.md)
    
- [GetNonIndexableItemDetails 操作](getnonindexableitemdetails-operation.md)
    
- [GetNonIndexableItemStatistics 操作](getnonindexableitemstatistics-operation.md)
    

