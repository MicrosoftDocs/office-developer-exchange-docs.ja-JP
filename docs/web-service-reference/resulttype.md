---
title: ResultType
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 488ee828-343f-4382-a5e8-eed1005f5dbc
description: ResultType 要素には、実行する検索の種類が含まれています。 検索の種類には、統計のみ、またはプレビューのみを指定できます。
ms.openlocfilehash: 6617c8b4b64cd9b6728317d7247bcc5378e488f0
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44465283"
---
# <a name="resulttype"></a>ResultType

**ResultType**要素には、実行する検索の種類が含まれています。 検索の種類には、統計のみ、またはプレビューのみを指定できます。 
  
```XML
<ResultType>StatisticsOnly | PreviewOnly</ResultType>
```

 **SearchResultType**
## <a name="attributes-and-elements"></a>属性と要素

以下のセクションで、属性、子要素、親要素について説明します。
  
### <a name="attributes"></a>属性

なし。
  
### <a name="child-elements"></a>子要素

なし。
  
### <a name="parent-elements"></a>親要素

[SearchMailboxesResult](searchmailboxesresult.md)  | [Searchmailboxes ボックス](searchmailboxes.md)
  
## <a name="text-value"></a>テキスト値

**ResultType**要素のテキスト値は、探索検索で返される結果の種類です。 **StatisticsOnly**のテキスト値は、検索統計を返します。 プレビュー**のみ**のテキスト値は、アイテムのプレビュー情報を返します。 
  
## <a name="remarks"></a>注釈

この要素は Exchange Server 2013 で導入されました。
  
この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。
  
## <a name="element-information"></a>要素の情報

|||
|:-----|:-----|
|Namespace  <br/> |https://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|スキーマ名  <br/> |メッセージスキーマ  <br/> |
|検証ファイル  <br/> |メッセージ .xsd  <br/> |
|空にすることができます。  <br/> |false  <br/> |
   

