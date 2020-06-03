---
title: CreateFolderResponse
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
api_name:
- CreateFolderResponse
api_type:
- schema
ms.assetid: 158adecc-491a-47d9-af73-acc2cd3f8566
description: CreateFolderResponse 要素は、CreateFolder 要求への応答を定義します。
ms.openlocfilehash: c51fb17f0b0f9c3dd1db8b0ff31f230fac37ce67
ms.sourcegitcommit: 88ec988f2bb67c1866d06b361615f3674a24e795
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44458881"
---
# <a name="createfolderresponse"></a><span data-ttu-id="fbc1f-103">CreateFolderResponse</span><span class="sxs-lookup"><span data-stu-id="fbc1f-103">CreateFolderResponse</span></span>

<span data-ttu-id="fbc1f-104">**Createfolderresponse**要素は、CreateFolder 要求への応答を定義します。</span><span class="sxs-lookup"><span data-stu-id="fbc1f-104">The **CreateFolderResponse** element defines a response to a CreateFolder request.</span></span> 
  
```xml
<CreateFolderResponse>
   <ResponseMessages/>
</CreateFolderResponse>
```

 <span data-ttu-id="fbc1f-105">**CreateFolderResponseType**</span><span class="sxs-lookup"><span data-stu-id="fbc1f-105">**CreateFolderResponseType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="fbc1f-106">属性と要素</span><span class="sxs-lookup"><span data-stu-id="fbc1f-106">Attributes and elements</span></span>

<span data-ttu-id="fbc1f-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="fbc1f-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="fbc1f-108">属性</span><span class="sxs-lookup"><span data-stu-id="fbc1f-108">Attributes</span></span>

<span data-ttu-id="fbc1f-109">なし。</span><span class="sxs-lookup"><span data-stu-id="fbc1f-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="fbc1f-110">子要素</span><span class="sxs-lookup"><span data-stu-id="fbc1f-110">Child elements</span></span>

|<span data-ttu-id="fbc1f-111">**Element**</span><span class="sxs-lookup"><span data-stu-id="fbc1f-111">**Element**</span></span>|<span data-ttu-id="fbc1f-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="fbc1f-112">**Description**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="fbc1f-113">ResponseMessages</span><span class="sxs-lookup"><span data-stu-id="fbc1f-113">ResponseMessages</span></span>](responsemessages.md) <br/> |<span data-ttu-id="fbc1f-114">Exchange Web サービス要求に対する応答メッセージを含みます。</span><span class="sxs-lookup"><span data-stu-id="fbc1f-114">Contains the response messages for an Exchange Web Services request.</span></span>  <br/> |
   
### <a name="parent-elements"></a><span data-ttu-id="fbc1f-115">親要素</span><span class="sxs-lookup"><span data-stu-id="fbc1f-115">Parent elements</span></span>

<span data-ttu-id="fbc1f-116">なし。</span><span class="sxs-lookup"><span data-stu-id="fbc1f-116">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="fbc1f-117">注釈</span><span class="sxs-lookup"><span data-stu-id="fbc1f-117">Remarks</span></span>

<span data-ttu-id="fbc1f-118">この要素を記述するスキーマは、クライアントアクセスサーバーの役割がインストールされている Microsoft Exchange Server 2007 を実行しているコンピューターの EWS 仮想ディレクトリにあります。</span><span class="sxs-lookup"><span data-stu-id="fbc1f-118">The schema that describes this element is located in the EWS virtual directory of the computer that is running MicrosoftExchange Server 2007 that has the Client Access server role installed.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="fbc1f-119">要素の情報</span><span class="sxs-lookup"><span data-stu-id="fbc1f-119">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="fbc1f-120">Namespace</span><span class="sxs-lookup"><span data-stu-id="fbc1f-120">Namespace</span></span>  <br/> |https://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="fbc1f-121">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="fbc1f-121">Schema name</span></span>  <br/> |<span data-ttu-id="fbc1f-122">メッセージスキーマ</span><span class="sxs-lookup"><span data-stu-id="fbc1f-122">Messages schema</span></span>  <br/> |
|<span data-ttu-id="fbc1f-123">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="fbc1f-123">Validation file</span></span>  <br/> |<span data-ttu-id="fbc1f-124">メッセージ .xsd</span><span class="sxs-lookup"><span data-stu-id="fbc1f-124">Messages.xsd</span></span>  <br/> |
|<span data-ttu-id="fbc1f-125">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="fbc1f-125">Can be empty</span></span>  <br/> |<span data-ttu-id="fbc1f-126">正しくない</span><span class="sxs-lookup"><span data-stu-id="fbc1f-126">False</span></span>  <br/> |
   
## <a name="see-also"></a><span data-ttu-id="fbc1f-127">関連項目</span><span class="sxs-lookup"><span data-stu-id="fbc1f-127">See also</span></span>



[<span data-ttu-id="fbc1f-128">CreateFolder 操作</span><span class="sxs-lookup"><span data-stu-id="fbc1f-128">CreateFolder operation</span></span>](createfolder-operation.md)
  
[<span data-ttu-id="fbc1f-129">CreateFolder</span><span class="sxs-lookup"><span data-stu-id="fbc1f-129">CreateFolder</span></span>](createfolder.md)


- [<span data-ttu-id="fbc1f-130">Exchange の EWS XML 要素</span><span class="sxs-lookup"><span data-stu-id="fbc1f-130">EWS XML elements in Exchange</span></span>](ews-xml-elements-in-exchange.md)

