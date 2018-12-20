---
title: Exchange において EWS を使用してメール メッセージに応答する
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
localization_priority: Normal
ms.assetid: 9d584991-4d67-4d36-ae2f-99970af8488f
description: Exchange で EWS マネージ API または EWS を使用してメール メッセージに応答する方法について説明します。
ms.openlocfilehash: 2f1428251084a7f2bf8d589a788c143f34b64d5c
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19759052"
---
# <a name="respond-to-email-messages-by-using-ews-in-exchange"></a>Exchange において EWS を使用してメール メッセージに応答する

Exchange で EWS マネージ API または EWS を使用してメール メッセージに応答する方法について説明します。
  
EWS マネージ API または EWS を使用して、メッセージに返信したり、受信者にメッセージに転送したりすることでメッセージに応答できます。
  
**表 1. メール メッセージに応答するための EWS マネージ API のメソッドと EWS 操作**

|**タスク**|**EWS マネージ API メソッド**|**EWS 操作**|
|:-----|:-----|:-----|
|メール メッセージに返信する  <br/> |[EmailMessage.Reply](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.reply%28v=exchg.80%29.aspx) <br/> [EmailMessage.CreateReply](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.createreply%28v=exchg.80%29.aspx) <br/> |[CreateItem](http://msdn.microsoft.com/library/fe6bb7fc-8918-4e6e-b0a1-b7e0ef44c3d1%28Office.15%29.aspx)。[Items](http://msdn.microsoft.com/library/d61ef1cc-ddfc-480a-9625-7b436cb33ae0%28Office.15%29.aspx) 要素には、[ReplyToItem](http://msdn.microsoft.com/library/35ee751a-41c0-4216-ad8b-78f7ada43a2f%28Office.15%29.aspx) または [ReplyAllToItem](http://msdn.microsoft.com/library/8ca970ca-ca73-40db-9233-7b271cc5f44f%28Office.15%29.aspx) という子要素が含まれます。  <br/> |
|メール メッセージの転送  <br/> |[EmailMessage.Forward](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.forward%28v=exchg.80%29.aspx) <br/> [EmailMessage.CreateForward](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.createforward%28v=exchg.80%29.aspx) <br/> |[CreateItem](http://msdn.microsoft.com/library/fe6bb7fc-8918-4e6e-b0a1-b7e0ef44c3d1%28Office.15%29.aspx)。[Items](http://msdn.microsoft.com/library/d61ef1cc-ddfc-480a-9625-7b436cb33ae0%28Office.15%29.aspx) 要素には [ForwardItem](http://msdn.microsoft.com/library/97786086-8b91-4471-8af8-d21e8d66de87%28Office.15%29.aspx) という子要素が含まれます。  <br/> |
   
## <a name="reply-to-an-email-message-by-using-the-ews-managed-api"></a>EWS マネージ API を使用してメール メッセージに返信する
<a name="bk_replyewsma"> </a>

EWS マネージ API には、メッセージへの応答に使用できる [ Reply ](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.reply%28v=exchg.80%29.aspx) と [CreateReply](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.createreply%28v=exchg.80%29.aspx) という 2 つのメソッドがあります。 **Reply** メソッドは、2 つのパラメーターだけを取ります。既存の本文の先頭に追加する応答メッセージと、すべての受信者に送信するか (true) 送信者だけに送信するか (false) を指定する **Boolean** 値です。 他の受信者をメッセージに追加したり、応答に追加プロパティを設定したり、添付物を追加したりする必要がある場合は、**CreateReply** メソッドを使用します。このメソッドでは、[EmailMessage](email-properties-and-elements-in-ews-in-exchange.md) オブジェクトで使用可能なすべての[ファースト クラスのプロパティ](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage%28v=exchg.80%29.aspx)を設定できます。 
  
次のコード例は、**Reply** メソッドを使用してメール メッセージに応答する方法を示しています。 
  
この例では、**service** が有効な [ExchangeService](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.exchangeservice%28v=exchg.80%29.aspx) オブジェクトであり、ユーザーが Exchange サーバーに既に認証されていると想定しています。 ローカル変数 *ItemId* は、対応するアイテムの [Id](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.item.id%28v=exchg.80%29.aspx) です。 この例では、[FindRecentlySent メソッド](#bk_findlast)を呼び出して、メッセージに返信済みのマークが付けられたことを確認します。 
  
```cs
// As a best practice, limit the properties returned by the Bind method to only those that are required.
PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.LastModifiedTime);
// Bind to the email message to reply to by using the ItemId.
// This method call results in a GetItem call to EWS.
EmailMessage message = EmailMessage.Bind(service, ItemId, propSet);
string myReply = "This is the message body of the email reply.";
bool replyToAll = false;
// Send the response message.
// This method call results in a CreateItem call to EWS.
message.Reply(myReply, replyToAll);
// Verify that the response was sent by calling FindRecentlySent.
FindRecentlySent(message);
```

次のコード例は、**CreateReply** メソッドを使用してメール メッセージに応答する方法を示しています。 
  
```cs
// Bind to the email message to reply to by using the ItemId.
// This method call results in a GetItem call to EWS.
EmailMessage message = EmailMessage.Bind(service, ItemId, BasePropertySet.IdOnly);
// Create the reply response message from the original email message.
// Indicate whether the message is a reply or reply all type of reply.
bool replyToAll = true;
ResponseMessage responseMessage = message.CreateReply(replyToAll);
// Prepend the reply to the message body. 
string myReply = "This is the message body of the email reply.";
responseMessage.BodyPrefix = myReply;
// Send the response message.
// This method call results in a CreateItem call to EWS.
responseMessage.SendAndSaveCopy();
// Check that the response was sent by calling FindRecentlySent.
FindRecentlySent(message);
```

応答メッセージに添付物を追加する場合は、**SendAndSaveCopy** メソッドへの呼び出しを次のコードに置換します。 
  
```cs
EmailMessage reply = responseMessage.Save();
reply.Attachments.AddFileAttachment("attachmentname.txt");
reply.Update(ConflictResolutionMode.AutoResolve);
reply.SendAndSaveCopy();
```

## <a name="reply-to-an-email-message-by-using-ews"></a>EWS を使用してメール メッセージに返信する
<a name="bk_replyews"> </a>

次のコード例は、EWS を使用してメッセージに返信する方法を示しています。 **MessageDisposition** 属性が **SendAndSaveCopy** に設定されている [CreateItem](http://msdn.microsoft.com/library/fe6bb7fc-8918-4e6e-b0a1-b7e0ef44c3d1%28Office.15%29.aspx) 操作を使用して、メッセージを送信し、送信済みアイテム フォルダーに応答を保存します。 [ReplyAllToItem](http://msdn.microsoft.com/library/8ca970ca-ca73-40db-9233-7b271cc5f44f%28Office.15%29.aspx) 要素を [Items](http://msdn.microsoft.com/library/d61ef1cc-ddfc-480a-9625-7b436cb33ae0%28Office.15%29.aspx) 要素の子として含めてメッセージ スレッド上の全員に返信するか、[ReplyToItem](http://msdn.microsoft.com/library/35ee751a-41c0-4216-ad8b-78f7ada43a2f%28Office.15%29.aspx) 要素を含めて送信者にのみ返信します。 
  
これは、[Reply](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.reply%28v=exchg.80%29.aspx) メソッドまたは [CreateReply](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.createreply%28v=exchg.80%29.aspx) メソッドを呼び出す際に、EWS マネージ API が送信する XML 要求でもあります。 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version=" Exchange2007_SP1" />
  </soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:ReplyAllToItem>
          <t:ReferenceItemId Id="AAMkADE4="
                             ChangeKey="CQAAABYA" />
          <t:NewBodyContent BodyType="HTML">This is the message body of the email reply.</t:NewBodyContent>
        </t:ReplyAllToItem>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>
```

サーバーは、**CreateItem** 要求に [CreateItemResponse](http://msdn.microsoft.com/library/742a46a0-2475-45a0-b44f-90639a3f5a43%28Office.15%29.aspx) メッセージで応答します。このメッセージには、応答が正常に作成および送信されたことを示す、**NoError** の [ResponseCode](http://msdn.microsoft.com/library/4b84d670-74c9-4d6d-84e7-f0a9f76f0d93%28Office.15%29.aspx) 要素の値が含まれます。
  
応答メッセージに添付物を追加する必要がある場合、前述したように **CreateItem** 操作を呼び出しますが、**MessageDisposition** は **SaveOnly** に変更します。 次に、[CreateAttachment](http://msdn.microsoft.com/library/e066db95-6963-4507-a8d0-8efad287f550%28Office.15%29.aspx) 操作を呼び出してから、[SendItem](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx) 操作を呼び出します。 
  
## <a name="forward-an-email-message-by-using-the-ews-managed-api"></a>EWS マネージ API を使用してメール メッセージを転送する
<a name="bk_forwardewsma"> </a>

EWS マネージ API には、メッセージの転送に使用できる [Forward](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.forward%28v=exchg.80%29.aspx) と [CreateForward](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.createforward%28v=exchg.80%29.aspx) という 2 つのメソッドがあります。 **Forward** メソッドは、2 つのパラメーターだけを取ります。使用するオーバーロードに応じて、既存の本文に先頭に付加するメッセージと、受信者の配列またはコレクションです。 転送するメッセージに添付物を追加したり、新しいメッセージで追加プロパティを設定したりする必要がある場合は、**CreateForward** メソッドを使用します。このメソッドでは、[EmailMessage](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage%28v=exchg.80%29.aspx) オブジェクトで使用可能なすべてのプロパティを設定できます。 
  
次のコード例は、**Forward** メソッドを使用してメール メッセージを 1 人の受信者に転送する方法を示しています。 
  
この例では、**service** が有効な [ExchangeService](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.exchangeservice%28v=exchg.80%29.aspx) オブジェクトであり、ユーザーが Exchange サーバーに既に認証されていると想定しています。 ローカル変数 *ItemId* は、転送するアイテムの [Id](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.item.id%28v=exchg.80%29.aspx) です。 この例では、[FindRecentlySent メソッド](#bk_findlast) を呼び出して、メッセージに転送済みのマークが付けられたことを確認します。 
  
```cs
// Bind to the email message to reply to by using the ItemId.
// This method call results in a GetItem call to EWS.
EmailMessage message = EmailMessage.Bind(service, ItemId, BasePropertySet.IdOnly);
string myForward = "This is the message body of the forwarded email.";
// Send the response message.
// This method call results in a CreateItem call to EWS.
message.Forward(myForward, "sadie@contoso.com");
// Verify that the forwarded response was sent by calling FindRecentlySent.
FindRecentlySent(message);
```

次のコード例は、**CreateForward** メソッドを使用してメール メッセージを 1 人の受信者に転送する方法を示しています。 
  
```cs
// Bind to the email message to reply to by using the ItemId.
// This method call results in a GetItem call to EWS.
EmailMessage message = EmailMessage.Bind(service, ItemId, BasePropertySet.IdOnly);
// Create the reply response message from the original email message.
// Indicate whether the message is a reply or reply all type of reply.
ResponseMessage forwardMessage = message.CreateForward();
// Set properties on the email message.
forwardMessage.ToRecipients.Add("sadie@contoso.com");
forwardMessage.Body = "Sadie,<br><br>I thought you'd be interested in this thread.<br><br>-Mack";
// Send and save a copy of the replied email message in the default Sent Items folder. 
forwardMessage.SendAndSaveCopy();
// Verify that the forwarded message was sent by calling FindRecentlySent.
FindRecentlySent(message);
```

転送されるメッセージに添付物を追加する場合は、**SendAndSaveCopy** メソッドへの呼び出しを次のコードに置換します。 
  
```cs
EmailMessage forward = forwardMessage.Save();
forward.Attachments.AddFileAttachment("attachmentname.txt");
forward.Update(ConflictResolutionMode.AutoResolve);
forward.SendAndSaveCopy();
```

## <a name="forward-an-email-message-by-using-ews"></a>EWS を使用してメール メッセージを転送する
<a name="bk_forwardews"> </a>

次のコード例は、EWS を使用してメッセージを転送する方法を示しています。 **MessageDisposition** 属性が **SendAndSaveCopy** に設定されている [CreateItem](http://msdn.microsoft.com/library/fe6bb7fc-8918-4e6e-b0a1-b7e0ef44c3d1%28Office.15%29.aspx) 操作を使用して、メッセージを送信し、送信済みアイテム フォルダーに応答を保存します。 [ForwardItem](http://msdn.microsoft.com/library/97786086-8b91-4471-8af8-d21e8d66de87%28Office.15%29.aspx) 要素は、そのアイテムが転送されるメッセージであることを示します。 
  
これは、[Forward](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.forward%28v=exchg.80%29.aspx) メソッドまたは [CreateForward](http://msdn.microsoft.com/ja-JP/library/microsoft.exchange.webservices.data.emailmessage.createforward%28v=exchg.80%29.aspx) メソッドを呼び出す際に、EWS マネージ API が送信する XML 要求でもあります。 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2007_SP1" />
  </soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:ForwardItem>
          <t:ToRecipients>
            <t:Mailbox>
              <t:EmailAddress>sadie@contoso.com</t:EmailAddress>
            </t:Mailbox>
          </t:ToRecipients>
          <t:ReferenceItemId Id="AAAMkADE="
                             ChangeKey="CQAAABYA" />
          <t:NewBodyContent BodyType="HTML">This is the message body of the forwarded email.</t:NewBodyContent>
        </t:ForwardItem>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>
```

サーバーは、**CreateItem** 要求に [CreateItemResponse](http://msdn.microsoft.com/library/742a46a0-2475-45a0-b44f-90639a3f5a43%28Office.15%29.aspx) メッセージで応答します。このメッセージには、転送済みメッセージが正常に作成および送信されたことを示す、**NoError** の [ResponseCode](http://msdn.microsoft.com/library/4b84d670-74c9-4d6d-84e7-f0a9f76f0d93%28Office.15%29.aspx) 要素の値が含まれます。
  
応答メッセージに添付物を追加する必要がある場合、**CreateItem** 操作を呼び出しますが、**MessageDisposition** を **SaveOnly** に変更します。 次に、[CreateAttachment](http://msdn.microsoft.com/library/e066db95-6963-4507-a8d0-8efad287f550%28Office.15%29.aspx) 操作を呼び出してから、[SendItem](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx) 操作を呼び出します。 
  
## <a name="find-the-message-last-replied-to-or-forwarded-by-using-the-ews-managed-api"></a>EWS マネージ API を使用して最後に返信または転送されたメッセージを検索する
<a name="bk_findlast"> </a>

次のコード例は、指定されたアイテムで実行された最後の動詞と、最後の動詞が実行された時間を検索する方法を示しています。このメソッドは、このトピックの他の EWS マネージ API コード例から呼び出され、返信または転送したアイテムが受信トレイで返信済みまたは転送済みとしてマークされていることを確認します。  
  
この例では、[PidTagLastVerbExecuted](http://msdn.microsoft.com/ja-JP/library/cc841968.aspx) (0x10820003) 拡張プロパティを使用して、メッセージが [返信]、[全員に返信]、[転送] のどれであるかを判別します。また、[PidTagLastVerbExecutionTime](http://msdn.microsoft.com/ja-JP/library/cc839918.aspx) (0x10820040) 拡張プロパティを使用して、返信または転送がいつ送信されたかを判別します。 
  
```cs
public static void FindRecentlySent(EmailMessage messageToCheck)
{
    // Create extended property definitions for PidTagLastVerbExecuted and PidTagLastVerbExecutionTime.
    ExtendedPropertyDefinition PidTagLastVerbExecuted = new ExtendedPropertyDefinition(0x1081, MapiPropertyType.Integer);
    ExtendedPropertyDefinition PidTagLastVerbExecutionTime = new ExtendedPropertyDefinition(0x1082, MapiPropertyType.SystemTime);
    PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, PidTagLastVerbExecutionTime, PidTagLastVerbExecuted);
    messageToCheck = EmailMessage.Bind(service, messageToCheck.Id, propSet);
    // Determine the last verb executed on the message and display output.
    object responseType;
    if (messageToCheck.TryGetProperty(PidTagLastVerbExecuted, out responseType))
    {
        object ReplyTime = null;
        switch (((Int32)responseType))
        {
            case 102: Console.WriteLine("A reply was sent to the '" + messageToCheck.Subject.ToString() + "' email message at");
                break;
            case 103: Console.WriteLine("A reply all was sent to the '" + messageToCheck.Subject.ToString() + "' email message at");
                break;
            case 104: Console.WriteLine("The '" + messageToCheck.Subject.ToString() + "' email message was forwarded at");
                break;
        }
        if (messageToCheck.TryGetProperty(PidTagLastVerbExecutionTime, out ReplyTime))
        {
            Console.WriteLine(((DateTime)ReplyTime).ToString() + ".");
        }
    }
    else
    {
        Console.WriteLine("No changes were made to  '" + messageToCheck.Subject.ToString() + "'.");
    }
}
```

## <a name="see-also"></a>関連項目


- [Exchange のメールと EWS](email-and-ews-in-exchange.md)
    
- [Exchange で EWS を使用してメール メッセージを送信する](how-to-send-email-messages-by-using-ews-in-exchange.md)
    

