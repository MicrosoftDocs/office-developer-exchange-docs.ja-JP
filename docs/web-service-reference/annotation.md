---
title: アノテーションフィーチャー
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: e0392635-9d0f-46d5-84ef-0a8a3036479a
description: Annotation 要素には、ユーザーが追加したオプションのメモが含まれています。
ms.openlocfilehash: 291d875085f9bc13e92a14b844b66878f5e6eb2a
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44466116"
---
# <a name="annotation"></a>アノテーションフィーチャー

**Annotation**要素には、ユーザーが追加したオプションのメモが含まれています。 
  
```XML
<Annotation></Annotation>
```

 **xs: 文字列**
## <a name="attributes-and-elements"></a>属性と要素

以下のセクションで、属性、子要素、親要素について説明します。
  
### <a name="attributes"></a>属性

なし。
  
### <a name="child-elements"></a>子要素

なし。
  
### <a name="parent-elements"></a>親要素

|**要素**|**説明**|
|:-----|:-----|
|[EnhancedLocation](enhancedlocation.md) <br/> |場所に関する名前、住所、オプションのメモなどの場所情報を指定します。  <br/> |
   
## <a name="text-value"></a>テキスト値

Annotation 要素のテキスト値は、ユーザーが場所に関するメモを追加したことを示します。
  
## <a name="remarks"></a>注釈

この要素は Exchange Server 2013 で導入されました。
  
この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。
  
## <a name="element-information"></a>要素の情報

|||
|:-----|:-----|
|Namespace  <br/> |https://schemas.microsoft.com/exchange/services/2006/types  <br/> |
|スキーマ名  <br/> |型スキーマ  <br/> |
|検証ファイル  <br/> |型 .xsd  <br/> |
|空にすることができます。  <br/> ||
   
## <a name="see-also"></a>関連項目

- [Exchange の EWS XML 要素](ews-xml-elements-in-exchange.md)

