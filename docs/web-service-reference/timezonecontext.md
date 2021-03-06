---
title: TimeZoneContext
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- TimeZoneContext
api_type:
- schema
ms.assetid: 573c462b-aa1d-4ba0-8852-e3f48b26873b
description: TimeZoneContext 要素は、簡易オブジェクトアクセスプロトコル (SOAP) ヘッダーで使用され、Exchange Web サービス (EWS) を使用して作成、更新、および取得されるオブジェクトの DateTime プロパティのタイムゾーンを割り当てるときに既定として使用されるタイムゾーン定義を指定します。
ms.openlocfilehash: 26727317ccf34338e8d62ec92bd7a44d43a6cfdb
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44460254"
---
# <a name="timezonecontext"></a>TimeZoneContext

**TimeZoneContext**要素は、簡易オブジェクトアクセスプロトコル (SOAP) ヘッダーで使用され、Exchange Web サービス (EWS) を使用して作成、更新、および取得されるオブジェクトの DateTime プロパティのタイムゾーンを割り当てるときに既定として使用されるタイムゾーン定義を指定します。 
  
```xml
<TimeZoneContext>
   <TimeZoneDefinition/>
</TimeZoneContext>
```

 **TimeZoneContextType**
## <a name="attributes-and-elements"></a>属性と要素

以下のセクションで、属性、子要素、親要素について説明します。
  
### <a name="attributes"></a>属性

なし。
  
### <a name="child-elements"></a>子要素

|**Element**|**説明**|
|:-----|:-----|
|[TimeZoneDefinition](timezonedefinition.md) <br/> |タイムゾーンを定義する期間と切り替えを指定します。  <br/> |
   
### <a name="parent-elements"></a>親要素

なし。
  
## <a name="remarks"></a>注釈

この要素を記述するスキーマは、クライアントアクセスサーバーの役割がインストールされている Microsoft Exchange Server を実行しているコンピューターの EWS 仮想ディレクトリにあります。
  
## <a name="element-information"></a>要素の情報

|||
|:-----|:-----|
|Namespace  <br/> |https://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|スキーマ名  <br/> |Types スキーマ  <br/> |
|検証ファイル  <br/> |型 .xsd  <br/> |
|空に設定可能  <br/> |正しくない  <br/> |
   
## <a name="see-also"></a>関連項目



- [Exchange の EWS XML 要素](ews-xml-elements-in-exchange.md)

