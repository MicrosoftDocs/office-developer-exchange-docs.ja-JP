---
title: End
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- End
api_type:
- schema
ms.assetid: 72329821-32ff-495d-b6e5-fdc011003c2e
description: End 要素は、期間の終わりを表します。
ms.openlocfilehash: d36f555d2ac9c0c1d82053029720ec17a53f2d92
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44456144"
---
# <a name="end"></a>End

**End**要素は、期間の終わりを表します。 
  
```xml
<End/>
```

 **DateTime**
## <a name="attributes-and-elements"></a>属性と要素

以下のセクションで、属性、子要素、親要素について説明します。
  
### <a name="attributes"></a>属性

なし。
  
### <a name="child-elements"></a>子要素

なし。
  
### <a name="parent-elements"></a>親要素

|**要素**|**説明**|
|:-----|:-----|
|[CalendarItem](calendaritem.md) <br/> |Exchange の予定表アイテムを表します。  <br/> |
|[FirstOccurrence](firstoccurrence.md) <br/> |定期的な予定表アイテムの最初の出現を表します。  <br/> |
|[LastOccurrence](lastoccurrence.md) <br/> |定期的な予定表アイテムの最後の発生を表します。  <br/> |
|[MeetingRequest](meetingrequest.md) <br/> |Exchange ストア内の会議出席依頼を表します。  <br/> |
|[発生](occurrence.md) <br/> |定期的な予定表アイテムの1つの変更されたアイテムを表します。  <br/> |
   
## <a name="text-value"></a>テキスト値

テキスト値は、期間の終わりを表します。
  
## <a name="remarks"></a>注釈

UpdateItem 操作は、Exchange ストアアイテムの[開始](start.md)時刻と**終了**時刻を設定できます。 UpdateItem 要求では、**終了**時刻を設定することなく、[開始](start.md)時刻を設定できます。 これにより、[開始](start.md)時刻が**終了**時刻よりも後の場合にエラーが発生することがあります。 時間を保持するためにその[開始](start.md)時刻が変更された場合、クライアントアプリケーションは**終了**時刻を調整する必要があることに注意してください。 
  
 **メモ**定期的なマスターアイテムの[開始](start.md)日と**終了**日に、最初に週単位の定期的なパターンと同じ日付が設定されていない場合は、タイムゾーンのオフセット情報は失われます。 
  
この要素を記述するスキーマは、Microsoft Exchange Server 2007 を実行しているコンピューターの EWS 仮想ディレクトリにあり、クライアントアクセスサーバーの役割がインストールされています。
  
## <a name="element-information"></a>要素の情報

|||
|:-----|:-----|
|Namespace  <br/> |https://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|スキーマ名  <br/> |Types スキーマ  <br/> |
|検証ファイル  <br/> |型 .xsd  <br/> |
|空に設定可能  <br/> |正しくない  <br/> |
   
## <a name="see-also"></a>関連項目



[WeeklyRecurrence](weeklyrecurrence.md)
  
 **End**


- [Exchange の EWS XML 要素](ews-xml-elements-in-exchange.md)

