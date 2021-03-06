---
title: FreeBusyViewType
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- FreeBusyViewType
api_type:
- schema
ms.assetid: 7c7f82ba-fa52-4a3e-bec7-39d373c66fc7
description: FreeBusyViewType 要素は、応答で返される空き時間情報の種類を表します。
ms.openlocfilehash: 3556ad236693ac9aa018b8aa3af7843765da6aa4
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44459568"
---
# <a name="freebusyviewtype"></a>FreeBusyViewType

**FreeBusyViewType**要素は、応答で返される空き時間情報の種類を表します。 
  
[GetUserAvailabilityResponse](getuseravailabilityresponse.md)
  
[FreeBusyResponseArray](freebusyresponsearray.md)
  
[FreeBusyResponse](freebusyresponse.md)
  
[FreeBusyView](freebusyview.md)
  
[FreeBusyViewType](freebusyviewtype.md)
  
```xml
<FreeBusyViewType>None or MergedOnly or FreeBusy or FreeBusyMerged or Detailed or DetailedMerged</FreeBusyViewType>
```

 **FreeBusyViewType**
## <a name="attributes-and-elements"></a>属性と要素

以下のセクションで、属性、子要素、親要素について説明します。
  
### <a name="attributes"></a>属性

なし。
  
### <a name="child-elements"></a>子要素

なし。
  
### <a name="parent-elements"></a>親要素

|**要素**|**説明**|
|:-----|:-----|
|[FreeBusyView](freebusyview.md) <br/> |特定のユーザーの空き時間情報が保存されています。  <br/> この要素の XPath 式を次に示します。  <br/>  `/GetUserAvailabilityResponse/FreeBusyResponseArray/FreeBusyResponse/FreeBusyView` <br/> |
   
## <a name="text-value"></a>テキスト値

テキスト値は必須です。 次の表に、この要素で使用できる値を示します。
  
|**値**|**説明**|
|:-----|:-----|
|なし  <br/> |この値は、要求に対しては無効です。 この値は、応答に対して有効です。  <br/> |
|MergedOnly  <br/> |集約された空き時間ストリームを表します。 1つのフォレストのターゲットユーザーが可用性サービスを構成していないフォレスト間のシナリオでは、リクエスターの可用性サービスが空き時間情報パブリックフォルダーからターゲットユーザーの空き時間情報を取得します。 パブリックフォルダーは、空き時間情報のみをマージされた形式で格納するため、 **MergedOnly**のみが利用可能な情報です。  <br/> |
|FreeBusy  <br/> |従来の状態の情報を表します。これは、空き時間、仮の予定、および OOF です。 これには、予定の開始時刻と終了時刻も含まれます。 このビューは、従来の空き時間表示よりも豊富です。個々の会議の開始時刻と終了時刻は、集計された空き時間ストリームの代わりに提供されます。  <br/> |
|FreeBusyMerged  <br/> |**FreeBusy**のすべてのプロパティを表し、マージされた空き時間情報を含むストリームを表します。  <br/> |
|詳細  <br/> |従来の状態の情報を表します。これは、空き時間、仮の予定、および OOF です。予定の開始/終了時刻。また、件名、場所、重要度など、予定のさまざまなプロパティを指定できます。 この要求されたビューは、要求元ユーザーが特権を持っている情報の最大量を返します。 マージされた空き時間情報のみを使用できる場合は、Microsoft Exchange Server 2003 フォレスト内のユーザーに関する情報を要求すると、 **MergedOnly**が返されます。 それ以外の場合は、 **FreeBusy**または**Detailed**が返されます。  <br/> 配布リストに**詳細**が指定されている場合は、リストのメンバーの空き時間情報が結合され、 **MergedOnly**が返されます。  <br/> |
|DetailedMerged  <br/> |マージされた空き時間情報のストリームに**詳細**が含まれるすべてのプロパティを表します。 マージされた空き時間情報のみが利用可能な場合 (たとえば、Exchange 2003 を実行しているコンピューター上にメールボックスが存在する場合)、 **MergedOnly**が返されます。 それ以外の場合は、 **FreeBusyMerged**または**DetailedMerged**が返されます。  <br/> |
   
## <a name="remarks"></a>注釈

この要素は、 [FreeBusyView](freebusyview.md)要素を使用する場合に必要です。 返される空き時間情報の種類は、 [Requestedview](requestedview.md)要素で指定されます。 この要素を記述するスキーマは、クライアントアクセスサーバーの役割がインストールされている Microsoft Exchange Server 2007 を実行しているコンピューターの EWS 仮想ディレクトリにあります。 
  
次の表は、さまざまなビューの種類と、それに対応する MAPI プロパティに対して返されるものを示しています。 各ビューの種類は、以前のビューの種類に基づいて作成されます。
  
|**FreeBusyViewType**|**Properties**|**MAPI Calendar プロパティ**|
|:-----|:-----|:-----|
|**MergedOnly** <br/> |MergedFreeBusyStream  <br/> ||
|**FreeBusy** <br/> |クラシック状態  <br/> |PropTag (0x80860003)  <br/> |
|**FreeBusy** <br/> |就業時間  <br/> ||
|**FreeBusy** <br/> |開始時刻  <br/> |PR_START_DATE  <br/> |
|**FreeBusy** <br/> |終了時刻  <br/> |PR_END_DATE  <br/> |
|**FreeBusyMerged** <br/> |クラシック状態  <br/> |PropTag (0x80860003)  <br/> |
|**FreeBusyMerged** <br/> |就業時間  <br/> ||
|**FreeBusyMerged** <br/> |開始時刻  <br/> |PR_START_DATE  <br/> |
|**FreeBusyMerged** <br/> |終了時刻  <br/> |PR_END_DATE  <br/> |
|**FreeBusyMerged** <br/> |MergedFreeBusyStream  <br/> ||
|**詳細** <br/> |クラシック状態  <br/> |PropTag (0x80860003)  <br/> |
|**詳細** <br/> |就業時間  <br/> ||
|**詳細** <br/> |開始時刻  <br/> |PR_START_DATE  <br/> |
|**詳細** <br/> |終了時刻  <br/> |PR_END_DATE  <br/> |
|**詳細** <br/> |件名  <br/> |PR_SUBJECT  <br/> |
|**詳細** <br/> |場所  <br/> |PR_LOCATION  <br/> |
|**詳細** <br/> |エントリ Id (プライベートでない場合)  <br/> ||
|**詳細** <br/> |プライベートフラグ  <br/> ||
|**詳細** <br/> |IsMeeting  <br/> ||
|**詳細** <br/> |IsRecurring  <br/> ||
|**詳細** <br/> |IsException  <br/> ||
|**詳細** <br/> |IsReminderSet  <br/> ||
|**詳細** <br/> |不在時のメッセージ (要求された場合)  <br/> ||
|**DetailedMerged** <br/> |クラシック状態  <br/> |PropTag (0x80860003)  <br/> |
|**DetailedMerged** <br/> |就業時間  <br/> ||
|**DetailedMerged** <br/> |開始時刻  <br/> |PR_START_DATE  <br/> |
|**DetailedMerged** <br/> |終了時刻  <br/> |PR_END_DATE  <br/> |
|**DetailedMerged** <br/> |件名  <br/> |PR_SUBJECT  <br/> |
|**DetailedMerged** <br/> |場所  <br/> |PR_LOCATION  <br/> |
|**DetailedMerged** <br/> |エントリ Id (プライベートでない場合)  <br/> ||
|**DetailedMerged** <br/> |プライベートフラグ  <br/> ||
|**DetailedMerged** <br/> |MergedFreeBusyStream  <br/> ||
|**DetailedMerged** <br/> |IsMeeting  <br/> ||
|**DetailedMerged** <br/> |IsRecurring  <br/> ||
|**DetailedMerged** <br/> |IsException  <br/> ||
|**DetailedMerged** <br/> |IsReminderSet  <br/> ||
   
## <a name="element-information"></a>要素の情報

|||
|:-----|:-----|
|Namespace  <br/> |https://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|スキーマ名  <br/> |Types スキーマ  <br/> |
|検証ファイル  <br/> |型 .xsd  <br/> |
|空に設定可能  <br/> |正しくない  <br/> |
   
## <a name="see-also"></a>関連項目



[GetUserAvailability 操作](getuseravailability-operation.md)
  
[GetUserAvailabilityResponse](getuseravailabilityresponse.md)


[ユーザーの空き時間情報の取得](https://msdn.microsoft.com/library/d4133fcb-9b0f-4e6b-aadf-a389da83516a%28Office.15%29.aspx)

