---
title: GetReminders 操作
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: overview
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 1b56f83f-3b87-4b55-8259-fde6692da681
description: GetReminders EWS 操作についての情報を検索します。
ms.openlocfilehash: dcbe20c674d7524a7776d374fa6964899abf472f
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44458307"
---
# <a name="getreminders-operation"></a><span data-ttu-id="8690e-103">GetReminders 操作</span><span class="sxs-lookup"><span data-stu-id="8690e-103">GetReminders operation</span></span>

<span data-ttu-id="8690e-104">**Getreminders** EWS 操作についての情報を検索します。</span><span class="sxs-lookup"><span data-stu-id="8690e-104">Find information about the **GetReminders** EWS operation.</span></span> 
  
<span data-ttu-id="8690e-105">**Getreminders** Exchange Web サービス (EWS) 操作は、予定表およびタスクアイテムのアラームを取得します。</span><span class="sxs-lookup"><span data-stu-id="8690e-105">The **GetReminders** Exchange Web Services (EWS) operation retrieves reminders for calendar and task items.</span></span> 
  
<span data-ttu-id="8690e-106">この操作は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="8690e-106">This operation was introduced in Exchange Server 2013.</span></span>
  
## <a name="using-the-getreminders-operation"></a><span data-ttu-id="8690e-107">GetReminders 操作の使用</span><span class="sxs-lookup"><span data-stu-id="8690e-107">Using the GetReminders operation</span></span>

<span data-ttu-id="8690e-108">**Getreminders**操作は、要求で渡された要素の値に応じて、ユーザーのメールボックス内の現在の予定表と今後の予定表およびタスクアイテムの事前通知を取得します。</span><span class="sxs-lookup"><span data-stu-id="8690e-108">The **GetReminders** operation gets reminders for current and future calendar and task items in the user's mailbox, depending on the element values passed in the request.</span></span> <span data-ttu-id="8690e-109">この操作では、アラームが設定されているタスクだけでなく、すべての現在の予定表および今後の予定表アイテムを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="8690e-109">The operation can retrieve all current and future calendar items as well as tasks that have a reminder set.</span></span> <span data-ttu-id="8690e-110">個人用の予定表アイテムは、応答に含まれます。</span><span class="sxs-lookup"><span data-stu-id="8690e-110">Private calendar items are included in responses.</span></span> <span data-ttu-id="8690e-111">アラームが設定されていないタスクは、応答には含まれません。また、アラームやフォローアップフラグが付いた電子メールには含まれません。</span><span class="sxs-lookup"><span data-stu-id="8690e-111">Tasks without reminders are not included in responses, nor are emails with reminders or follow up flags.</span></span> 
  
<span data-ttu-id="8690e-112">現在のすべてのアラームを取得するには、 [ReminderType](remindertype.md)を [**すべて**] に設定し、 [EndTime](endtime-remindermessagedatatype.md)を現在の時刻に設定することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="8690e-112">To retrieve all current reminders, we recommend setting the [ReminderType](remindertype.md) to **All** and the [EndTime](endtime-remindermessagedatatype.md) to the current time.</span></span> 
  
<span data-ttu-id="8690e-113">要求に[Begintime](begintime.md)および**EndTime**要素が含まれている場合、応答には、その間に発生する予定表およびタスクアイテムのアラームが含まれています。これらのアイテムの間には、 **begintime**と**endtime**の間に発生するアラームがあります。</span><span class="sxs-lookup"><span data-stu-id="8690e-113">If the [BeginTime](begintime.md) and **EndTime** elements are included in the request, the response includes reminders for any calendar and task items that occur between have a reminder that occurs between the **BeginTime** and **EndTime**.</span></span>
  
<span data-ttu-id="8690e-114">次の表では、 **Begintime**要素と**EndTime**要素が含まれている場合の**ReminderType**要素の動作について説明します。</span><span class="sxs-lookup"><span data-stu-id="8690e-114">The following table describes the behavior of the **ReminderType** element when the **BeginTime** and **EndTime** elements are included.</span></span> 
  
|<span data-ttu-id="8690e-115">ReminderType \* \* 要素の値 \* \*</span><span class="sxs-lookup"><span data-stu-id="8690e-115">\*\*\*\*ReminderType\*\* element value\*\*</span></span>|<span data-ttu-id="8690e-116">**説明**</span><span class="sxs-lookup"><span data-stu-id="8690e-116">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="8690e-117">すべて</span><span class="sxs-lookup"><span data-stu-id="8690e-117">All</span></span>  <br/> |<span data-ttu-id="8690e-118">**Begintime**と**EndTime**の間に発生するアラーム。</span><span class="sxs-lookup"><span data-stu-id="8690e-118">Reminders that occur between the **BeginTime** and **EndTime**.</span></span>  <br/> |
|<span data-ttu-id="8690e-119">Current</span><span class="sxs-lookup"><span data-stu-id="8690e-119">Current</span></span>  <br/> |<span data-ttu-id="8690e-120">**すべて**の返信と、要求された時間枠より前のアラームに加えて、保存期間に関係なくすべての予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="8690e-120">Reminders returned by **All**, plus reminders that are earlier than the requested time window if the event is still ongoing, plus all appointments regardless of age.</span></span>  <br/> |
|<span data-ttu-id="8690e-121">古く</span><span class="sxs-lookup"><span data-stu-id="8690e-121">Old</span></span>  <br/> |<span data-ttu-id="8690e-122">完了していない**すべて**のマイナスイベントによって返されるアラーム。すべての予定を差し引いた。</span><span class="sxs-lookup"><span data-stu-id="8690e-122">Reminders returned by **All**, minus events that haven't completed yet, minus all appointments.</span></span> <span data-ttu-id="8690e-123">**前**の値を使用するには、 **Begintime**要素と**EndTime**要素を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8690e-123">The **BeginTime** and **EndTime** elements must be set to use the **Old** value.</span></span>  <br/> |
   
### <a name="getreminders-operation-soap-headers"></a><span data-ttu-id="8690e-124">GetReminders 操作の SOAP ヘッダー</span><span class="sxs-lookup"><span data-stu-id="8690e-124">GetReminders operation SOAP headers</span></span>

<span data-ttu-id="8690e-125">**Getreminders**操作では、次の表に記載されている SOAP ヘッダーを使用できます。</span><span class="sxs-lookup"><span data-stu-id="8690e-125">The **GetReminders** operation can use the SOAP headers that are listed in the following table.</span></span> 
  
|<span data-ttu-id="8690e-126">**ヘッダー名**</span><span class="sxs-lookup"><span data-stu-id="8690e-126">**Header name**</span></span>|<span data-ttu-id="8690e-127">**要素**</span><span class="sxs-lookup"><span data-stu-id="8690e-127">**Element**</span></span>|<span data-ttu-id="8690e-128">**説明**</span><span class="sxs-lookup"><span data-stu-id="8690e-128">**Description**</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="8690e-129">**偽装**</span><span class="sxs-lookup"><span data-stu-id="8690e-129">**Impersonation**</span></span> <br/> |[<span data-ttu-id="8690e-130">ExchangeImpersonation</span><span class="sxs-lookup"><span data-stu-id="8690e-130">ExchangeImpersonation</span></span>](exchangeimpersonation.md) <br/> |<span data-ttu-id="8690e-131">クライアントアプリケーションが偽装しているユーザーを識別します。</span><span class="sxs-lookup"><span data-stu-id="8690e-131">Identifies the user whom the client application is impersonating.</span></span> <span data-ttu-id="8690e-132">このヘッダーは、要求に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8690e-132">This header is applicable to a request.</span></span>  <br/> |
|<span data-ttu-id="8690e-133">**MailboxCulture**</span><span class="sxs-lookup"><span data-stu-id="8690e-133">**MailboxCulture**</span></span> <br/> |[<span data-ttu-id="8690e-134">MailboxCulture</span><span class="sxs-lookup"><span data-stu-id="8690e-134">MailboxCulture</span></span>](mailboxculture.md) <br/> |<span data-ttu-id="8690e-135">RFC 3066 で定義されているように、メールボックスへのアクセスに使用されるカルチャ (言語の識別用のタグ) を識別します。</span><span class="sxs-lookup"><span data-stu-id="8690e-135">Identifies the culture, as defined in RFC 3066, "Tags for the Identification of Languages", to be used to access the mailbox.</span></span> <span data-ttu-id="8690e-136">このヘッダーは、要求に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8690e-136">This header is applicable to a request.</span></span>  <br/> |
|<span data-ttu-id="8690e-137">**RequestVersion**</span><span class="sxs-lookup"><span data-stu-id="8690e-137">**RequestVersion**</span></span> <br/> |[<span data-ttu-id="8690e-138">RequestServerVersion</span><span class="sxs-lookup"><span data-stu-id="8690e-138">RequestServerVersion</span></span>](requestserverversion.md) <br/> |<span data-ttu-id="8690e-139">操作要求のスキーマバージョンを識別します。</span><span class="sxs-lookup"><span data-stu-id="8690e-139">Identifies the schema version for the operation request.</span></span> <span data-ttu-id="8690e-140">このヘッダーは、要求に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8690e-140">This header is applicable to a request.</span></span>  <br/> |
|<span data-ttu-id="8690e-141">**ServerVersion**</span><span class="sxs-lookup"><span data-stu-id="8690e-141">**ServerVersion**</span></span> <br/> |[<span data-ttu-id="8690e-142">ServerVersionInfo</span><span class="sxs-lookup"><span data-stu-id="8690e-142">ServerVersionInfo</span></span>](serverversioninfo.md) <br/> |<span data-ttu-id="8690e-143">要求に応答したサーバーのバージョンを識別します。</span><span class="sxs-lookup"><span data-stu-id="8690e-143">Identifies the version of the server that responded to the request.</span></span> <span data-ttu-id="8690e-144">このヘッダーは応答に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8690e-144">This header is applicable to a response.</span></span>  <br/> |
   
## <a name="getreminders-operation-request-example"></a><span data-ttu-id="8690e-145">GetReminders 操作の要求の例</span><span class="sxs-lookup"><span data-stu-id="8690e-145">GetReminders operation request example</span></span>

<span data-ttu-id="8690e-146">次の**Getreminders**操作要求の例は、 **Begintime**と**EndTime**の間に発生する最初の5つの予定表アイテムを取得する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8690e-146">The following example of a **GetReminders** operation request shows how to retrieve the first five calendar items that occur between the **BeginTime** and **EndTime**.</span></span>
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="https://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:GetReminders>
      <m:EndTime>2014-04-16T21:00:00Z</m:EndTime>
      <m:ReminderType>All</m:ReminderType>
    </m:GetReminders>
  </soap:Body>
</soap:Envelope>
```

<span data-ttu-id="8690e-147">要求 SOAP 本文の例には、次の要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8690e-147">The example request SOAP body contains the following elements:</span></span>
  
- [<span data-ttu-id="8690e-148">GetReminders</span><span class="sxs-lookup"><span data-stu-id="8690e-148">GetReminders</span></span>](getreminders.md)
    
- [<span data-ttu-id="8690e-149">EndTime</span><span class="sxs-lookup"><span data-stu-id="8690e-149">EndTime</span></span>](endtime-remindermessagedatatype.md)
    
- [<span data-ttu-id="8690e-150">ReminderType</span><span class="sxs-lookup"><span data-stu-id="8690e-150">ReminderType</span></span>](remindertype.md)
    
<span data-ttu-id="8690e-151">SOAP 本文には、次の要素も含めることができます。</span><span class="sxs-lookup"><span data-stu-id="8690e-151">The SOAP body can also contain the following elements:</span></span>
  
- [<span data-ttu-id="8690e-152">BeginTime</span><span class="sxs-lookup"><span data-stu-id="8690e-152">BeginTime</span></span>](begintime.md)
    
- [<span data-ttu-id="8690e-153">MaxItems</span><span class="sxs-lookup"><span data-stu-id="8690e-153">MaxItems</span></span>](maxitems.md)
    
## <a name="successful-getreminders-operation-response"></a><span data-ttu-id="8690e-154">成功した GetReminders 操作の応答</span><span class="sxs-lookup"><span data-stu-id="8690e-154">Successful GetReminders operation response</span></span>

<span data-ttu-id="8690e-155">次の例は、 **Getreminders**操作要求に対する正常な応答を示しています。</span><span class="sxs-lookup"><span data-stu-id="8690e-155">The following example shows a successful response to a **GetReminders** operation request.</span></span> <span data-ttu-id="8690e-156">応答には、"teams meeting" 予定表アイテムの事前通知と、"会議メモを送信するタスク" タスクに関する事前通知が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8690e-156">The response contains a reminder for the "Team meeting" calendar item and a reminder for the "Task to send meeting notes" task.</span></span> 
  
> [!NOTE]
> <span data-ttu-id="8690e-157">読みやすくするために、識別子が短縮されています。</span><span class="sxs-lookup"><span data-stu-id="8690e-157">Identifiers have been shortened to preserve readability.</span></span> 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <ServerVersionInfo MajorVersion="15"
                       MinorVersion="0"
                       MajorBuildNumber="918"
                       MinorBuildNumber="7"
                       Version="V2_10"
                       xmlns:h="https://schemas.microsoft.com/exchange/services/2006/types"
                       xmlns="https://schemas.microsoft.com/exchange/services/2006/types"
                       xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                       xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </s:Header>
  <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <GetRemindersResponse ResponseClass="Success"
                          xmlns="https://schemas.microsoft.com/exchange/services/2006/messages">
      <ResponseCode>NoError</ResponseCode>
      <Reminders>
        <Reminder xmlns="https://schemas.microsoft.com/exchange/services/2006/types">
          <Subject>Team meeting</Subject>
          <Location />
          <ReminderTime>2014-04-15T21:00:00Z</ReminderTime>
          <StartDate>2014-04-15T21:00:00Z</StartDate>
          <EndDate>2014-04-15T21:30:00Z</EndDate>
          <ItemId Id="vQAAAA=="
                  ChangeKey="DwAAABYAAAB4to43JyybTYwHLBM1k8MxAAACRoV4" />
          <RecurringMasterItemId Id="K7u5AAA=" ChangeKey="DwAAABYAAAB4to43JyybTYwHLBM1k8MxAAACRoV0" />
          <ReminderGroup>Calendar</ReminderGroup>
          <UID>6CF2FA62</UID>
        </Reminder>
        <Reminder xmlns="https://schemas.microsoft.com/exchange/services/2006/types">
          <Subject>Task to send meeting notes</Subject>
          <Location />
          <ReminderTime>2014-04-16T14:00:00Z</ReminderTime>
          <StartDate>0001-01-02T00:00:00Z</StartDate>
          <EndDate>0001-01-02T00:00:00Z</EndDate>
          <ItemId Id="vAAAAA=="
                  ChangeKey="EwAAABQAAACOs0HEMq1WTKpI7sNu5qXNAAAIDg==" />
          <ReminderGroup>Task</ReminderGroup>
          <UID>vAAAAA==</UID>
        </Reminder>
      </Reminders>
    </GetRemindersResponse>
  </s:Body>
</s:Envelope>
```

<span data-ttu-id="8690e-158">応答 SOAP 本文には、次の要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8690e-158">The response SOAP body contains the following elements:</span></span>
  
- [<span data-ttu-id="8690e-159">GetRemindersResponse</span><span class="sxs-lookup"><span data-stu-id="8690e-159">GetRemindersResponse</span></span>](getremindersresponse.md)
    
- [<span data-ttu-id="8690e-160">Reminders</span><span class="sxs-lookup"><span data-stu-id="8690e-160">Reminders</span></span>](reminders.md)
    
- [<span data-ttu-id="8690e-161">Reminder</span><span class="sxs-lookup"><span data-stu-id="8690e-161">Reminder</span></span>](reminder.md)
    
- [<span data-ttu-id="8690e-162">件名</span><span class="sxs-lookup"><span data-stu-id="8690e-162">Subject</span></span>](subject.md)
    
- [<span data-ttu-id="8690e-163">場所</span><span class="sxs-lookup"><span data-stu-id="8690e-163">Location</span></span>](location-remindermessagedatatype.md)
    
- [<span data-ttu-id="8690e-164">ReminderTime</span><span class="sxs-lookup"><span data-stu-id="8690e-164">ReminderTime</span></span>](remindertime.md)
    
- [<span data-ttu-id="8690e-165">StartDate</span><span class="sxs-lookup"><span data-stu-id="8690e-165">StartDate</span></span>](startdate.md)
    
- [<span data-ttu-id="8690e-166">EndDate</span><span class="sxs-lookup"><span data-stu-id="8690e-166">EndDate</span></span>](enddate-remindertype.md)
    
- [<span data-ttu-id="8690e-167">ItemId</span><span class="sxs-lookup"><span data-stu-id="8690e-167">ItemId</span></span>](itemid.md)
    
- [<span data-ttu-id="8690e-168">RecurringMasterItemId</span><span class="sxs-lookup"><span data-stu-id="8690e-168">RecurringMasterItemId</span></span>](recurringmasteritemid.md)
    
- [<span data-ttu-id="8690e-169">ReminderGroup</span><span class="sxs-lookup"><span data-stu-id="8690e-169">ReminderGroup</span></span>](remindergroup.md)
    
- [<span data-ttu-id="8690e-170">UID</span><span class="sxs-lookup"><span data-stu-id="8690e-170">UID</span></span>](uid-remindertype.md)
    
## <a name="getreminders-operation-error-response-example"></a><span data-ttu-id="8690e-171">GetReminders 操作エラー応答の例</span><span class="sxs-lookup"><span data-stu-id="8690e-171">GetReminders operation error response example</span></span>

<span data-ttu-id="8690e-172">次の例は、 **Getreminders**操作要求へのエラー応答を示しています。</span><span class="sxs-lookup"><span data-stu-id="8690e-172">The following example shows an error response to a **GetReminders** operation request.</span></span> <span data-ttu-id="8690e-173">これは、終了日が開始日よりも前の要求に対する応答です。</span><span class="sxs-lookup"><span data-stu-id="8690e-173">This is a response to a request in which the end date was earlier than the start date.</span></span> 
  
```XML
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
  <s:Header>
    <ServerVersionInfo MajorVersion="15"
                       MinorVersion="0"
                       MajorBuildNumber="918"
                       MinorBuildNumber="7"
                       Version="V2_10"
                       xmlns:h="https://schemas.microsoft.com/exchange/services/2006/types"
                       xmlns="https://schemas.microsoft.com/exchange/services/2006/types"
                       xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                       xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" />
  </s:Header>
  <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <GetRemindersResponse ResponseClass="Error"
                          xmlns="https://schemas.microsoft.com/exchange/services/2006/messages">
      <MessageText>EndDate is earlier than StartDate</MessageText>
      <ResponseCode>ErrorInvalidOperation</ResponseCode>
      <DescriptiveLinkKey>0</DescriptiveLinkKey>
    </GetRemindersResponse>
  </s:Body>
</s:Envelope>
```

<span data-ttu-id="8690e-174">エラー応答 SOAP 本文には、次の要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8690e-174">The error response SOAP body contains the following elements:</span></span>
  
- [<span data-ttu-id="8690e-175">GetRemindersResponse</span><span class="sxs-lookup"><span data-stu-id="8690e-175">GetRemindersResponse</span></span>](getremindersresponse.md)
    
- [<span data-ttu-id="8690e-176">MessageText</span><span class="sxs-lookup"><span data-stu-id="8690e-176">MessageText</span></span>](messagetext.md)
    
- [<span data-ttu-id="8690e-177">ResponseCode</span><span class="sxs-lookup"><span data-stu-id="8690e-177">ResponseCode</span></span>](responsecode.md)
    
- [<span data-ttu-id="8690e-178">DescriptiveLinkKey</span><span class="sxs-lookup"><span data-stu-id="8690e-178">DescriptiveLinkKey</span></span>](descriptivelinkkey.md)
    
<span data-ttu-id="8690e-179">EWS で汎用的でこの操作に固有の追加のエラーコードについては、「応答」[を参照してください。](responsecode.md)</span><span class="sxs-lookup"><span data-stu-id="8690e-179">For additional error codes that are generic to EWS and specific to this operation, see [ResponseCode](responsecode.md).</span></span>
  
## <a name="see-also"></a><span data-ttu-id="8690e-180">関連項目</span><span class="sxs-lookup"><span data-stu-id="8690e-180">See also</span></span>


- [<span data-ttu-id="8690e-181">PerformReminderAction</span><span class="sxs-lookup"><span data-stu-id="8690e-181">PerformReminderAction</span></span>](performreminderaction.md)
    

