---
title: 示
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- TimeStamp
api_type:
- schema
ms.assetid: 5eae859a-5a74-4bf6-b196-d1b2fd38501a
description: Timestamp 要素は、メールボックスイベントのタイムスタンプを表します。
ms.openlocfilehash: f2280d4eab67b603963c4f0a7468bf35a2b63a88
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44459890"
---
# <a name="timestamp"></a>示

**Timestamp**要素は、メールボックスイベントのタイムスタンプを表します。 
  
```xml
<TimeStamp/>
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
|[CopiedEvent](copiedevent.md) <br/> |アイテムまたはフォルダーがコピーされるイベントを表します。  <br/> |
|[対する createdevent](createdevent.md) <br/> |アイテムまたはフォルダーが作成されるイベントを表します。  <br/> |
|[DeletedEvent](deletedevent.md) <br/> |アイテムまたはフォルダーが削除されるイベントを表します。  <br/> |
|[対する modifiedevent](modifiedevent.md) <br/> |アイテムまたはフォルダーが変更されるイベントを表します。  <br/> |
|[MovedEvent](movedevent.md) <br/> |ある親フォルダーから別の親フォルダーにアイテムまたはフォルダーを移動するイベントを表します。  <br/> |
|[NewMailEvent](newmailevent.md) <br/> |メールボックス内の新しいメールアイテムによってトリガーされるイベントを表します。  <br/> |
   
## <a name="text-value"></a>テキスト値

このプロパティに値を設定するには、 SetExtrusionDirection メソッドを適用します。
  
## <a name="remarks"></a>注釈

この要素は、主に、イベント頻度をクライアントに決定するために使用できます。 これは[Statusevent](statusevent.md)では存在しません。
  
この要素を記述するスキーマは、クライアントアクセスサーバーの役割がインストールされている Microsoft Exchange Server 2007 を実行しているコンピューターの EWS 仮想ディレクトリにあります。
  
## <a name="element-information"></a>要素の情報

|||
|:-----|:-----|
|Namespace  <br/> |https://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|スキーマ名  <br/> |Types スキーマ  <br/> |
|検証ファイル  <br/> |型 .xsd  <br/> |
|空にすることができます。  <br/> |正しくない  <br/> |
   
## <a name="see-also"></a>関連項目



[サブスクライブ操作](subscribe-operation.md)
  
[GetEvents 操作](getevents-operation.md)
  
[Unsubscribe 操作](unsubscribe-operation.md)

