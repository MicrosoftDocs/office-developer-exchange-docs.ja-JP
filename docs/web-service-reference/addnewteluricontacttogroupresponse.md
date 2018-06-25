---
title: AddNewTelUriContactToGroupResponse
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: abff2306-a3a7-489a-b548-2edbc1eb5cc4
description: AddNewTelUriContactToGroupResponse 要素は、AddNewTelUriContactToGroup の WSDL 操作の結果のデータを指定します。
ms.openlocfilehash: ddf038af2f4dec8cff98c4453e867af1fe7b8076
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19759296"
---
# <a name="addnewteluricontacttogroupresponse"></a><span data-ttu-id="41224-103">AddNewTelUriContactToGroupResponse</span><span class="sxs-lookup"><span data-stu-id="41224-103">AddNewTelUriContactToGroupResponse</span></span>

<span data-ttu-id="41224-104">**AddNewTelUriContactToGroupResponse**要素は、 **AddNewTelUriContactToGroup**の WSDL 操作の結果のデータを指定します。</span><span class="sxs-lookup"><span data-stu-id="41224-104">The **AddNewTelUriContactToGroupResponse** element specifies the result data for the **AddNewTelUriContactToGroup** WSDL operation.</span></span> 
  
```XML
<AddNewTelUriContactToGroupResponse>
   <Persona/>
   <MessageText/>
   <ResponseCode/>
   <DescriptiveLinkKey/>
</AddNewTelUriContactToGroupResponse>
```

 <span data-ttu-id="41224-105">**AddNewTelUriContactToGroupResponseMessageType**</span><span class="sxs-lookup"><span data-stu-id="41224-105">**AddNewTelUriContactToGroupResponseMessageType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="41224-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="41224-106">Attributes and elements</span></span>

<span data-ttu-id="41224-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="41224-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="41224-108">属性</span><span class="sxs-lookup"><span data-stu-id="41224-108">Attributes</span></span>

<span data-ttu-id="41224-109">なし。</span><span class="sxs-lookup"><span data-stu-id="41224-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="41224-110">子要素</span><span class="sxs-lookup"><span data-stu-id="41224-110">Child elements</span></span>

<span data-ttu-id="41224-111">[ペルソナ](persona.md) | [メッセージ テキスト](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md)</span><span class="sxs-lookup"><span data-stu-id="41224-111">[Persona](persona.md) | [MessageText](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md)</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="41224-112">親要素</span><span class="sxs-lookup"><span data-stu-id="41224-112">Parent elements</span></span>

<span data-ttu-id="41224-113">なし。</span><span class="sxs-lookup"><span data-stu-id="41224-113">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="41224-114">備考</span><span class="sxs-lookup"><span data-stu-id="41224-114">Remarks</span></span>

<span data-ttu-id="41224-115">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="41224-115">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="41224-116">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="41224-116">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="41224-117">要素情報</span><span class="sxs-lookup"><span data-stu-id="41224-117">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="41224-118">名前空間</span><span class="sxs-lookup"><span data-stu-id="41224-118">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="41224-119">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="41224-119">Schema name</span></span>  <br/> |<span data-ttu-id="41224-120">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="41224-120">Messages schema</span></span>  <br/> |
|<span data-ttu-id="41224-121">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="41224-121">Validation file</span></span>  <br/> |<span data-ttu-id="41224-122">Messages.xsd</span><span class="sxs-lookup"><span data-stu-id="41224-122">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="41224-123">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="41224-123">Can be empty</span></span>  <br/> |<span data-ttu-id="41224-124">False</span><span class="sxs-lookup"><span data-stu-id="41224-124">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="41224-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="41224-125">See also</span></span>

- [<span data-ttu-id="41224-126">AddNewTelUriContactToGroup 操作</span><span class="sxs-lookup"><span data-stu-id="41224-126">AddNewTelUriContactToGroup operation</span></span>](addnewteluricontacttogroup-operation.md)
- [<span data-ttu-id="41224-127">Exchange での EWS の XML 要素</span><span class="sxs-lookup"><span data-stu-id="41224-127">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)

