---
title: IconIndex
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 92020822-2a86-4dfc-aee1-3067af4d4edf
description: IconIndex 要素は、アイテムまたは会話のアイコンのインデックスを識別します。
ms.openlocfilehash: 8ada9da2df1cf128009c9b153736b434925f1a3f
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19831848"
---
# <a name="iconindex"></a><span data-ttu-id="d856f-103">IconIndex</span><span class="sxs-lookup"><span data-stu-id="d856f-103">IconIndex</span></span>

<span data-ttu-id="d856f-104">**IconIndex**要素は、アイテムまたは会話のアイコンのインデックスを識別します。</span><span class="sxs-lookup"><span data-stu-id="d856f-104">The **IconIndex** element identifies the icon index for an item or conversation.</span></span> 
  
```XML
<IconIndex>Default | PostItem | MailRead | MailUnread | MailReplied | MailForwarded | MailEncrypted | MailSmimeSigned | MailEncrytedReplied | MailSmimeSignedReplied | MailEncryptedForwarded | MailSmimeSignedForwarded | MailEncryptedRead | MailSmimeSignedRead | MailIrm | MailIrmForwarded | MailIrmReplied | SmsSubmitted | SmsRoutedToDeliveryPoint | SmsRoutedToExternalMessagingSystem | SmsDelivered | OutlookDefaultForContacts | AppointmentItem | AppointmentRecur | AppointmentMeet | AppointmentMeetRecur | AppointmentMeetNY | AppointmentMeetYes | AppointmentMeetNo | AppointmentMeetMaybe | AppointmentMeetCancel | AppointmentMeetInfo | TaskItem | TaskRecur | TaskOwned | TaskDelegated</IconIndex>
```

 <span data-ttu-id="d856f-105">**IconIndexType**</span><span class="sxs-lookup"><span data-stu-id="d856f-105">**IconIndexType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="d856f-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="d856f-106">Attributes and elements</span></span>

<span data-ttu-id="d856f-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="d856f-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="d856f-108">属性</span><span class="sxs-lookup"><span data-stu-id="d856f-108">Attributes</span></span>

<span data-ttu-id="d856f-109">なし。</span><span class="sxs-lookup"><span data-stu-id="d856f-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="d856f-110">子要素</span><span class="sxs-lookup"><span data-stu-id="d856f-110">Child elements</span></span>

<span data-ttu-id="d856f-111">なし。</span><span class="sxs-lookup"><span data-stu-id="d856f-111">None.</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="d856f-112">親要素</span><span class="sxs-lookup"><span data-stu-id="d856f-112">Parent elements</span></span>

<span data-ttu-id="d856f-113">[会話 (ConversationType)](conversation-conversationtype.md) | [アイテム](item.md) | [連絡先](contact.md) | [DistributionList](distributionlist.md) | [メッセージ](message-ex15websvcsotherref.md) | [カレンダー項目](calendaritem.md) | [PostItem](postitem.md) | [タスク](task.md)</span><span class="sxs-lookup"><span data-stu-id="d856f-113">[Conversation (ConversationType)](conversation-conversationtype.md) | [Item](item.md) | [Contact](contact.md) | [DistributionList](distributionlist.md) | [Message](message-ex15websvcsotherref.md) | [CalendarItem](calendaritem.md) | [PostItem](postitem.md) | [Task](task.md)</span></span>
  
## <a name="text-value"></a><span data-ttu-id="d856f-114">テキスト値</span><span class="sxs-lookup"><span data-stu-id="d856f-114">Text value</span></span>

<span data-ttu-id="d856f-115">次の表には、 **IconIndex**要素の可能なテキスト値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="d856f-115">The following table contains the possible text values for the **IconIndex** element.</span></span> 
  
|<span data-ttu-id="d856f-116">**値**</span><span class="sxs-lookup"><span data-stu-id="d856f-116">**Value**</span></span>|<span data-ttu-id="d856f-117">**説明**</span><span class="sxs-lookup"><span data-stu-id="d856f-117">**Description**</span></span>|
|:-----|:-----|
|||
|<span data-ttu-id="d856f-118">Default</span><span class="sxs-lookup"><span data-stu-id="d856f-118">Default</span></span>  <br/> |<span data-ttu-id="d856f-119">既定のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-119">Specifies the default icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-120">PostItem</span><span class="sxs-lookup"><span data-stu-id="d856f-120">PostItem</span></span>  <br/> |<span data-ttu-id="d856f-121">投稿アイテムのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-121">Specifies the icon for a post item.</span></span>  <br/> |
|<span data-ttu-id="d856f-122">MailRead</span><span class="sxs-lookup"><span data-stu-id="d856f-122">MailRead</span></span>  <br/> |<span data-ttu-id="d856f-123">開封済みのメールを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-123">Specifies the mail read icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-124">MailUnread</span><span class="sxs-lookup"><span data-stu-id="d856f-124">MailUnread</span></span>  <br/> |<span data-ttu-id="d856f-125">未読のメールのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-125">Specifies the unread mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-126">MailReplied</span><span class="sxs-lookup"><span data-stu-id="d856f-126">MailReplied</span></span>  <br/> |<span data-ttu-id="d856f-127">[メール] アイコンに、返信を指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-127">Specifies the replied to mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-128">MailForwarded</span><span class="sxs-lookup"><span data-stu-id="d856f-128">MailForwarded</span></span>  <br/> |<span data-ttu-id="d856f-129">転送されたメールのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-129">Specifies the forwarded mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-130">MailEncrypted</span><span class="sxs-lookup"><span data-stu-id="d856f-130">MailEncrypted</span></span>  <br/> |<span data-ttu-id="d856f-131">暗号化されたメール] アイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-131">Specifies the encrypted mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-132">MailSmimeSigned</span><span class="sxs-lookup"><span data-stu-id="d856f-132">MailSmimeSigned</span></span>  <br/> |<span data-ttu-id="d856f-133">/Multipurpose をセキュリティで保護されたインターネット メール拡張 (S/MIME) 署名されたメールのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-133">Specifies the Secure/Multipurpose Internet Mail Extensions (S/MIME) signed mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-134">MailEncryptedReplied</span><span class="sxs-lookup"><span data-stu-id="d856f-134">MailEncryptedReplied</span></span>  <br/> |<span data-ttu-id="d856f-135">[メール] アイコンを暗号化された応答を指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-135">Specifies the encrypted replied to mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-136">MailSmimeSignedReplied</span><span class="sxs-lookup"><span data-stu-id="d856f-136">MailSmimeSignedReplied</span></span>  <br/> |<span data-ttu-id="d856f-137">秒/MIME 署名返信メールのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-137">Specifies the S/MIME signed replied to mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-138">MailEncryptedForwarded</span><span class="sxs-lookup"><span data-stu-id="d856f-138">MailEncryptedForwarded</span></span>  <br/> |<span data-ttu-id="d856f-139">暗号化された転送されたメール] アイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-139">Specifies the encrypted forwarded mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-140">MailSmimeSignedForwarded</span><span class="sxs-lookup"><span data-stu-id="d856f-140">MailSmimeSignedForwarded</span></span>  <br/> |<span data-ttu-id="d856f-141">S/MIME 署名されたメール] アイコンを転送するかを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-141">Specifies the S/MIME signed forwarded mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-142">MailEncryptedRead</span><span class="sxs-lookup"><span data-stu-id="d856f-142">MailEncryptedRead</span></span>  <br/> |<span data-ttu-id="d856f-143">暗号化された読み取りメール] アイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-143">Specifies the encrypted read mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-144">MailSmimeSignedRead</span><span class="sxs-lookup"><span data-stu-id="d856f-144">MailSmimeSignedRead</span></span>  <br/> |<span data-ttu-id="d856f-145">メール開封済みの S/MIME 署名を指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-145">Specifies the S/MIME signed read mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-146">MailIrm</span><span class="sxs-lookup"><span data-stu-id="d856f-146">MailIrm</span></span>  <br/> |<span data-ttu-id="d856f-147">情報権利管理の IRM で保護されたメール] アイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-147">Specifies the Information Rights Management (IRM)-protected mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-148">MailIrmForwarded</span><span class="sxs-lookup"><span data-stu-id="d856f-148">MailIrmForwarded</span></span>  <br/> |<span data-ttu-id="d856f-149">[メール] アイコンを転送された IRM で保護されたを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-149">Specifies the IRM-protected forwarded mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-150">MailIrmReplied</span><span class="sxs-lookup"><span data-stu-id="d856f-150">MailIrmReplied</span></span>  <br/> |<span data-ttu-id="d856f-151">返信 IRM で保護されたメール] アイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-151">Specifies the IRM-protected replied to mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-152">SmsSubmitted</span><span class="sxs-lookup"><span data-stu-id="d856f-152">SmsSubmitted</span></span>  <br/> |<span data-ttu-id="d856f-153">ショート メッセージ サービス (SMS) のルーティングのために送信] アイコンのメールを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-153">Specifies the icon mail submitted for Short Message Service (SMS) routing.</span></span>  <br/> |
|<span data-ttu-id="d856f-154">SmsRoutedToDeliveryPoint</span><span class="sxs-lookup"><span data-stu-id="d856f-154">SmsRoutedToDeliveryPoint</span></span>  <br/> |<span data-ttu-id="d856f-155">SMS が外部の配信ポイントにルーティングするためのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-155">Specifies the icon for SMS routing to an external delivery point.</span></span>  <br/> |
|<span data-ttu-id="d856f-156">SmsRoutedToExternalMessagingSystem</span><span class="sxs-lookup"><span data-stu-id="d856f-156">SmsRoutedToExternalMessagingSystem</span></span>  <br/> |<span data-ttu-id="d856f-157">SMS が外部のメッセージング システムにルーティングするためのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-157">Specifies the icon for SMS routing to an external messaging system.</span></span>  <br/> |
|<span data-ttu-id="d856f-158">SmsDelivered</span><span class="sxs-lookup"><span data-stu-id="d856f-158">SmsDelivered</span></span>  <br/> |<span data-ttu-id="d856f-159">SMS 配信されたメールのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-159">Specifies the SMS delivered mail icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-160">OutlookDefaultForContacts</span><span class="sxs-lookup"><span data-stu-id="d856f-160">OutlookDefaultForContacts</span></span>  <br/> |<span data-ttu-id="d856f-161">連絡先の既定のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-161">Specifies the default icon for contacts.</span></span>  <br/> |
|<span data-ttu-id="d856f-162">AppointmentItem</span><span class="sxs-lookup"><span data-stu-id="d856f-162">AppointmentItem</span></span>  <br/> |<span data-ttu-id="d856f-163">予定項目のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-163">Specifies the appointment item icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-164">AppointmentRecur</span><span class="sxs-lookup"><span data-stu-id="d856f-164">AppointmentRecur</span></span>  <br/> |<span data-ttu-id="d856f-165">定期的な予定のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-165">Specifies the recurring appointment icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-166">AppointmentMeet</span><span class="sxs-lookup"><span data-stu-id="d856f-166">AppointmentMeet</span></span>  <br/> |<span data-ttu-id="d856f-167">会議アイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-167">Specifies the meeting icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-168">AppointmentMeetRecur</span><span class="sxs-lookup"><span data-stu-id="d856f-168">AppointmentMeetRecur</span></span>  <br/> |<span data-ttu-id="d856f-169">定期的な会議のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-169">Specifies the recurring meeting icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-170">AppointmentMeetNY</span><span class="sxs-lookup"><span data-stu-id="d856f-170">AppointmentMeetNY</span></span>  <br/> |<span data-ttu-id="d856f-171">会議への暫定的な応答のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-171">Specifies the icon for a tentative response to the meeting.</span></span>  <br/> |
|<span data-ttu-id="d856f-172">AppointmentMeetYes</span><span class="sxs-lookup"><span data-stu-id="d856f-172">AppointmentMeetYes</span></span>  <br/> |<span data-ttu-id="d856f-173">会議の指定承認アイコンです。</span><span class="sxs-lookup"><span data-stu-id="d856f-173">Specifies the meeting acceptance icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-174">AppointmentMeetNo</span><span class="sxs-lookup"><span data-stu-id="d856f-174">AppointmentMeetNo</span></span>  <br/> |<span data-ttu-id="d856f-175">会議の辞退したアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-175">Specifies the meeting declined icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-176">AppointmentMeetMaybe</span><span class="sxs-lookup"><span data-stu-id="d856f-176">AppointmentMeetMaybe</span></span>  <br/> |<span data-ttu-id="d856f-177">会議にも応答のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-177">Specifies the icon for a maybe response to the meeting.</span></span>  <br/> |
|<span data-ttu-id="d856f-178">AppointmentMeetCancel</span><span class="sxs-lookup"><span data-stu-id="d856f-178">AppointmentMeetCancel</span></span>  <br/> |<span data-ttu-id="d856f-179">会議キャンセルのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-179">Specifies the meeting cancel icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-180">AppointmentMeetInfo</span><span class="sxs-lookup"><span data-stu-id="d856f-180">AppointmentMeetInfo</span></span>  <br/> |<span data-ttu-id="d856f-181">会議を指定する情報のアイコンです。</span><span class="sxs-lookup"><span data-stu-id="d856f-181">Specifies the meeting information icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-182">TaskItem</span><span class="sxs-lookup"><span data-stu-id="d856f-182">TaskItem</span></span>  <br/> |<span data-ttu-id="d856f-183">作業項目のアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-183">Specifies the task item icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-184">TaskRecur</span><span class="sxs-lookup"><span data-stu-id="d856f-184">TaskRecur</span></span>  <br/> |<span data-ttu-id="d856f-185">定期的なタスクのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-185">Specifies the recurring task icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-186">TaskOwned</span><span class="sxs-lookup"><span data-stu-id="d856f-186">TaskOwned</span></span>  <br/> |<span data-ttu-id="d856f-187">アイコンを所有するタスクを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-187">Specifies the task owned icon.</span></span>  <br/> |
|<span data-ttu-id="d856f-188">TaskDelegated</span><span class="sxs-lookup"><span data-stu-id="d856f-188">TaskDelegated</span></span>  <br/> |<span data-ttu-id="d856f-189">委任されたタスクのアイコンを指定します。</span><span class="sxs-lookup"><span data-stu-id="d856f-189">Specifies the task delegated icon.</span></span>  <br/> |
   
## <a name="remarks"></a><span data-ttu-id="d856f-190">備考</span><span class="sxs-lookup"><span data-stu-id="d856f-190">Remarks</span></span>

<span data-ttu-id="d856f-191">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="d856f-191">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="d856f-192">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="d856f-192">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="d856f-193">要素情報</span><span class="sxs-lookup"><span data-stu-id="d856f-193">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="d856f-194">名前空間</span><span class="sxs-lookup"><span data-stu-id="d856f-194">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|<span data-ttu-id="d856f-195">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="d856f-195">Schema name</span></span>  <br/> |<span data-ttu-id="d856f-196">タイプのスキーマ</span><span class="sxs-lookup"><span data-stu-id="d856f-196">Types schema</span></span>  <br/> |
|<span data-ttu-id="d856f-197">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="d856f-197">Validation file</span></span>  <br/> |<span data-ttu-id="d856f-198">Types.xsd</span><span class="sxs-lookup"><span data-stu-id="d856f-198">Types.xsd</span></span>  <br/> |
|<span data-ttu-id="d856f-199">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="d856f-199">Can be empty</span></span>  <br/> |<span data-ttu-id="d856f-200">false</span><span class="sxs-lookup"><span data-stu-id="d856f-200">false</span></span>  <br/> |
   
