---
title: Exchange EWS を使用して予定および会議を取得します。
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
localization_priority: Normal
ms.assetid: 1bae582a-8cb3-4e77-be2a-7e107fad26fe
description: EWS マネージ API または Exchange の EWS を使用して、予定と会議を取得する方法を説明します。
ms.openlocfilehash: 0f5eb135142e807f30f48f01d7948fbdbf147ac2
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/11/2018
ms.locfileid: "19758950"
---
# <a name="get-appointments-and-meetings-by-using-ews-in-exchange"></a>Exchange EWS を使用して予定および会議を取得します。

EWS マネージ API または Exchange の EWS を使用して、予定と会議を取得する方法を説明します。
  
[CalendarFolder.FindAppointments](http://msdn.microsoft.com/en-us/library/dd636179%28v=exchg.80%29.aspx) EWS のマネージ API のメソッドまたは[FindItem](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx) EWS の操作を使用して、予定表フォルダーから予定と会議を取得できます。 
  
## <a name="get-appointments-by-using-the-ews-managed-api"></a>EWS マネージ API を使用して、予定を取得する
<a name="bk_retrieveappsEWSMA"> </a>

次のコード例は、EWS マネージ API を使用して、指定した開始時刻から終了時刻までの範囲にあるユーザーの予定を取得する方法を示しています。
  
```cs
       // Initialize values for the start and end times, and the number of appointments to retrieve.
            DateTime startDate = DateTime.Now;
            DateTime endDate = startDate.AddDays(30);
            const int NUM_APPTS = 5;
            // Initialize the calendar folder object with only the folder ID. 
            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
            // Set the start and end time and number of appointments to retrieve.
            CalendarView cView = new CalendarView(startDate, endDate, NUM_APPTS);
            // Limit the properties returned to the appointment's subject, start time, and end time.
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);
            // Retrieve a collection of appointments by using the calendar view.
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);
            Console.WriteLine("\nThe first " + NUM_APPTS + " appointments on your calendar from " + startDate.Date.ToShortDateString() + 
                              " to " + endDate.Date.ToShortDateString() + " are: \n");
            
            foreach (Appointment a in appointments)
            {
                Console.Write("Subject: " + a.Subject.ToString() + " ");
                Console.Write("Start: " + a.Start.ToString() + " ");
                Console.Write("End: " + a.End.ToString());
                Console.WriteLine();
            }

```

以下に示すのは、コード例からの出力です。
  
2013/9/20 2013/8/21 から予定表の最初の 5 つの予定は次のとおりです。 
  
件名: contoso 社の開発者のチーム会議の開始: 2013/8/21 午後 12時 30分: 00 終了: 2013/8/21 午後 1時 00分: 00
  
件名: 日常会議開始: 2013/8/21 午後 1時 00分: 00 終了: 2013/8/21 午後 2時 00分: 00
  
件名: 営業チームの開始とランチ: 2013/8/21 午後 2時 30分: 00 終了: 2013/8/21 午後 3時 30分: 00
  
件名: クラブの開始時テニス: 2013/8/22 午前 11時 00分: 00 終了: 2013/8/22 12時 00分: 00
  
件名: オンライン トレーニングの web キャスト: 2013/8/22 午後 2時 00分: 00 終了: 2013/8/22 午後 3時 00分: 00
## <a name="get-appointments-by-using-ews"></a>EWS を使用して、予定を取得します 
<a name="bk_xml"> </a>

[FindItem](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)操作でフォルダーの ID を返す[GetFolder](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)操作要求を次の XML に表示されます。 
  
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
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="calendar" />
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
```

次の XML は、 **GetFolder**の応答を示しています。 **フォルダー Id**と**変更キー**の属性が読みやすくするために短縮されますに注意してください。 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <h:ServerVersionInfo MajorVersion="15" MinorVersion="0" MajorBuildNumber="731" MinorBuildNumber="10" Version="V2_3" 
 xmlns:h="http://schemas.microsoft.com/exchange/services/2006/types" 
 xmlns="http://schemas.microsoft.com/exchange/services/2006/types" 
 xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </s:Header>
  <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <m:GetFolderResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
 xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseMessages>
        <m:GetFolderResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:Folders>
            <t:CalendarFolder>
              <t:FolderId Id="AAMk" ChangeKey="AgAA" />
            </t:CalendarFolder>
          </m:Folders>
        </m:GetFolderResponseMessage>
      </m:ResponseMessages>
    </m:GetFolderResponse>
  </s:Body>
</s:Envelope>
```

次の XML は、要求された予定を返すために使用する**FindItem**要求を示しています。 **フォルダー Id**と**変更キー**の属性が読みやすくするために短縮されますに注意してください。 
  
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
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="calendar:Start" />
          <t:FieldURI FieldURI="calendar:End" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView MaxEntriesReturned="5" StartDate="2013-08-21T17:30:24.127Z" EndDate="2013-09-20T17:30:24.127Z" />
      <m:ParentFolderIds>
        <t:FolderId Id="AAMk" ChangeKey="AgAA" />
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>
```

**FindItem**応答を次の XML に表示されます。 **アイテム Id**と**変更キー**の属性が読みやすくするために短縮されますに注意してください。 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <h:ServerVersionInfo MajorVersion="15" MinorVersion="0" MajorBuildNumber="731" MinorBuildNumber="10" Version="V2_3" 
 xmlns:h="http://schemas.microsoft.com/exchange/services/2006/types" 
 xmlns="http://schemas.microsoft.com/exchange/services/2006/types" 
 xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </s:Header>
  <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <m:FindItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <m:ResponseMessages>
        <m:FindItemResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:RootFolder TotalItemsInView="33" IncludesLastItemInRange="false">
            <t:Items>
              <t:CalendarItem>
                <t:ItemId Id="AAMk" ChangeKey="DwAA" />
                <t:Subject>Contoso devs team meeting</t:Subject>
                <t:Start>2013-08-21T19:30:00Z</t:Start>
                <t:End>2013-08-21T20:00:00Z</t:End>
              </t:CalendarItem>
              <t:CalendarItem>
                <t:ItemId Id="AAMk" ChangeKey="DwAA" />
                <t:Subject>Daily status meeting</t:Subject>
                <t:Start>2013-08-21T20:00:00Z</t:Start>
                <t:End>2013-08-21T21:00:00Z</t:End>
              </t:CalendarItem>
              <t:CalendarItem>
                <t:ItemId Id="AAMk" ChangeKey="DwAA" />
                <t:Subject>Lunch with sales team</t:Subject>
                <t:Start>2013-08-21T21:30:00Z</t:Start>
                <t:End>2013-08-21T22:30:00Z</t:End>
              </t:CalendarItem>
              <t:CalendarItem>
                <t:ItemId Id="AAMk" ChangeKey="DwAA" />
                <t:Subject>Tennis at the club</t:Subject>
                <t:Start>2013-08-22T18:00:00Z</t:Start>
                <t:End>2013-08-22T19:00:00Z</t:End>
              </t:CalendarItem>
              <t:CalendarItem>
                <t:ItemId Id="AAMkA" ChangeKey="DwAA" />
                <t:Subject>Online training webcast</t:Subject>
                <t:Start>2013-08-22T21:00:00Z</t:Start>
                <t:End>2013-08-22T22:00:00Z</t:End>
              </t:CalendarItem>
            </t:Items>
          </m:RootFolder>
        </m:FindItemResponseMessage>
      </m:ResponseMessages>
    </m:FindItemResponse>
  </s:Body>
</s:Envelope>
```

## <a name="recurring-meetings-and-the-calendar-view"></a>定期的な会議および予定表の表示
<a name="bk_recurring"> </a>

一連の定期的な一連の定期的な例外の発生は、メールボックス内の実際の項目ではないではなく、定期的なマスターへの添付ファイルとして内部に格納されているために、予定表フォルダーがメールボックス内の他のフォルダーとは少し異なります。 つまり、**開始**の間の値を返す EWS 要求を作成することができ、 **FindItems**の EWS のマネージ API のいずれかを使用して、**最後**の値が[ExchangeService.FindItems](http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.exchangeservice.finditems%28v=exchg.80%29.aspx)または EWS[などのメソッドをオーバー ロードFindItem](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)操作、例外および出現を検索するすべての予定表アイテムの添付ファイル テーブルを EWS は表示されません。 
  
代わりに、どのような本当に実行するは、*データ ビュー*の 2 つの SQL テーブルでは、共用体に[CalendarView](http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.calendarview%28v=exchg.80%29.aspx)オブジェクトを使用して適用することに類似のもの。 パフォーマンス上の理由から、お勧めの予定や会議を取得するだけでなく、必要な特定のプロパティの数を示すことにより、応答のサイズを制限する[プロパティ設定](http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.propertyset%28v=exchg.80%29.aspx)のプロパティを使用することに注意してください。 
  
## <a name="see-also"></a>関連項目
<a name="bk_additional"> </a>

- [Calendars and EWS in Exchange](calendars-and-ews-in-exchange.md)
    
- [Exchange 2013 の EWS を使用して予定および会議を作成します。](how-to-create-appointments-and-meetings-by-using-ews-in-exchange-2013.md)
    
- [Exchange EWS を使用して予定および会議を更新します。](how-to-update-appointments-and-meetings-by-using-ews-in-exchange.md)
    
- [予定を削除して、Exchange で EWS を使用して会議をキャンセル](how-to-delete-appointments-and-cancel-meetings-by-using-ews-in-exchange.md)
    
- [Exchange の Web サービス クライアントを開発する](develop-web-service-clients-for-exchange.md)
    
