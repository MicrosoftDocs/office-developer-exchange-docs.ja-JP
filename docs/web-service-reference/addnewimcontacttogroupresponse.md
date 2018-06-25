---
title: AddNewImContactToGroupResponse
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: e1a1c5fa-9e1e-4ee7-bb19-77c29b47ecbb
description: AddNewImContactToGroupResponse 要素は、AddNewImContactToGroup 要求への応答を定義します。
ms.openlocfilehash: dedd5587fd4deef1287b377084c1a4bcfc91e80c
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19759291"
---
# <a name="addnewimcontacttogroupresponse"></a><span data-ttu-id="09fa6-103">AddNewImContactToGroupResponse</span><span class="sxs-lookup"><span data-stu-id="09fa6-103">AddNewImContactToGroupResponse</span></span>

<span data-ttu-id="09fa6-104">**AddNewImContactToGroupResponse**要素は、 **AddNewImContactToGroup**要求への応答を定義します。</span><span class="sxs-lookup"><span data-stu-id="09fa6-104">The **AddNewImContactToGroupResponse** element defines a response to an **AddNewImContactToGroup** request.</span></span> 
  
```XML
<AddNewImContactToGroupResponse>
   <MessageText/>
   <ResponseCode/>
   <DescriptiveLinkKey/>
   <MessageXml/>
   <Persona/>
</AddNewImContactToGroupResponse>
```

 <span data-ttu-id="09fa6-105">**AddNewImContactToGroupResponseMessageType**</span><span class="sxs-lookup"><span data-stu-id="09fa6-105">**AddNewImContactToGroupResponseMessageType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="09fa6-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="09fa6-106">Attributes and elements</span></span>

<span data-ttu-id="09fa6-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="09fa6-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="09fa6-108">属性</span><span class="sxs-lookup"><span data-stu-id="09fa6-108">Attributes</span></span>

<span data-ttu-id="09fa6-109">なし。</span><span class="sxs-lookup"><span data-stu-id="09fa6-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="09fa6-110">子要素</span><span class="sxs-lookup"><span data-stu-id="09fa6-110">Child elements</span></span>

<span data-ttu-id="09fa6-111">[メッセージ テキスト](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md) | [MessageXml](messagexml.md) | [ペルソナ](persona.md)</span><span class="sxs-lookup"><span data-stu-id="09fa6-111">[MessageText](messagetext.md) | [ResponseCode](responsecode.md) | [DescriptiveLinkKey](descriptivelinkkey.md) | [MessageXml](messagexml.md) | [Persona](persona.md)</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="09fa6-112">親要素</span><span class="sxs-lookup"><span data-stu-id="09fa6-112">Parent elements</span></span>

<span data-ttu-id="09fa6-113">なし。</span><span class="sxs-lookup"><span data-stu-id="09fa6-113">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="09fa6-114">備考</span><span class="sxs-lookup"><span data-stu-id="09fa6-114">Remarks</span></span>

<span data-ttu-id="09fa6-115">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="09fa6-115">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="09fa6-116">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="09fa6-116">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="09fa6-117">要素情報</span><span class="sxs-lookup"><span data-stu-id="09fa6-117">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="09fa6-118">名前空間</span><span class="sxs-lookup"><span data-stu-id="09fa6-118">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="09fa6-119">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="09fa6-119">Schema name</span></span>  <br/> |<span data-ttu-id="09fa6-120">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="09fa6-120">Messages schema</span></span>  <br/> |
|<span data-ttu-id="09fa6-121">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="09fa6-121">Validation file</span></span>  <br/> |<span data-ttu-id="09fa6-122">messages.xsd</span><span class="sxs-lookup"><span data-stu-id="09fa6-122">messages.xsd</span></span>  <br/> |
|<span data-ttu-id="09fa6-123">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="09fa6-123">Can be empty</span></span>  <br/> |<span data-ttu-id="09fa6-124">false</span><span class="sxs-lookup"><span data-stu-id="09fa6-124">false</span></span>  <br/> |
   

