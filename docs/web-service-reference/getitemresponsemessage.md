---
title: GetItemResponseMessage
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- GetItemResponseMessage
api_type:
- schema
ms.assetid: cc583723-54d1-4a17-8c1f-6586f70fdefd
description: GetItemResponseMessage 要素には、1つの GetItem 操作要求の状態と結果が含まれています。
ms.openlocfilehash: bd25a82641259e1546bad6eef5c2f6f8f03e98cb
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44458671"
---
# <a name="getitemresponsemessage"></a>GetItemResponseMessage

**GetItemResponseMessage**要素には、1つの[GetItem 操作](getitem-operation.md)要求の状態と結果が含まれています。 
  
- [GetItemResponse](getitemresponse.md) 
- [ResponseMessages](responsemessages.md)
- [GetItemResponseMessage](getitemresponsemessage.md)
  
```xml
<GetItemResponseMessage ResponseClass="">
   <MessageText/>
   <ResponseCode/>
   <DescriptiveLinkKey/>
   <MessageXml/>
   <Items/>
</GetItemResponseMessage>
```

**ItemInfoResponseMessageType**

## <a name="attributes-and-elements"></a>属性と要素

以下のセクションで、属性、子要素、親要素について説明します。
  
### <a name="attributes"></a>属性

|**属性**|**説明**|
|:-----|:-----|
|**ResponseClass** <br/> | [GetItem 操作](getitem-operation.md)の応答の状態を表します。 <br/><br/>この属性には、次の値が有効です。<br/><br/>-成功<br/>-Warning<br/>-エラー |
   
#### <a name="responseclass-attribute-values"></a>ResponseClass 属性の値

|**値**|**説明**|
|:-----|:-----|
|**Success** <br/> |満たされる要求を記述します。  <br/> |
|**Warning** <br/> | 処理されなかった要求を示します。 要求内のアイテムが処理され、その後のアイテムを処理できなかったときにエラーが発生した場合は、警告が返されることがあります。<br/><br/>次に、警告のソースの例を示します。<br/><br/>-バッチ処理中に Exchange ストアがオフラインになります。<br/>-Active Directory ドメインサービス (AD DS) がオフラインになっています。<br/>-メールボックスは移動されます。<br/>-MDB がオフラインになっています。<br/>-パスワードの有効期限が切れています。  <br/>-クォータを超過しました。 |
|**Error** <br/> | 満たされない要求を記述します。<br/><br/>エラーのソースの例を次に示します。<br/><br/>-無効な属性または要素<br/>-属性または要素が範囲外です<br/>-不明なタグ<br/>-属性または要素がコンテキスト内で有効ではありません<br/>-クライアントによる権限のないアクセスの試行<br/>-有効なクライアント側の呼び出しに応答した場合のサーバー側エラー<br/><br/>エラーに関する情報については、応答[secの](responsecode.md)要素と[messagetext](messagetext.md)要素を参照してください。 |
   
### <a name="child-elements"></a>子要素

|**Element**|**説明**|
|:-----|:-----|
|[MessageText](messagetext.md) <br/> |応答の状態を説明するテキストを提供します。  <br/> |
|[ResponseCode](responsecode.md) <br/> |要求で発生した特定のエラーを識別するエラーコードを提供します。  <br/> |
|[DescriptiveLinkKey](descriptivelinkkey.md) <br/> |現在未使用で、今後の使用のために予約されています。 このプロパティには0の値が含まれています。  <br/> |
|[MessageXml](messagexml.md) <br/> |エラー応答に関する追加情報を提供します。  <br/> |
|[Items](items.md) <br/> |返されるアイテムの配列を格納します。  <br/> |
   
### <a name="parent-elements"></a>親要素

|**要素**|**説明**|
|:-----|:-----|
|[ResponseMessages](responsemessages.md) <br/> |Exchange Web サービス要求に対する応答メッセージを含みます。  <br/> |
   
## <a name="remarks"></a>注釈

この要素を記述するスキーマは、Microsoft Exchange Server 2010 を実行しているコンピューターの EWS 仮想ディレクトリにあり、クライアントアクセスサーバーの役割がインストールされています。
  
## <a name="element-information"></a>要素の情報

|||
|:-----|:-----|
|Namespace  <br/> |https://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|スキーマ名  <br/> |メッセージスキーマ  <br/> |
|検証ファイル  <br/> |メッセージ .xsd  <br/> |
|空に設定可能  <br/> |正しくない  <br/> |
   
## <a name="see-also"></a>関連項目

- [GetItem](getitem.md)
- [GetItem 操作](getitem-operation.md)

