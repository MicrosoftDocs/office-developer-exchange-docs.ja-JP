---
title: GetUserPhoto
manager: sethgros
ms.date: 09/17/2015
ms.audience: Developer
ms.topic: reference
ms.prod: office-online-server
localization_priority: Normal
ms.assetid: 0e524a09-86f8-4a71-ac5c-66527ae70790
description: GetUserPhoto 要素には、ユーザーの写真を取得する要求が含まれています。
ms.openlocfilehash: 1fa1db6d903f83ad54ffadb188ba79ee178de797
ms.sourcegitcommit: 34041125dc8c5f993b21cebfc4f8b72f0fd2cb6f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2018
ms.locfileid: "19831695"
---
# <a name="getuserphoto"></a><span data-ttu-id="c3a7b-103">GetUserPhoto</span><span class="sxs-lookup"><span data-stu-id="c3a7b-103">GetUserPhoto</span></span>

<span data-ttu-id="c3a7b-104">**GetUserPhoto**要素には、ユーザーの写真を取得する要求が含まれています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-104">The **GetUserPhoto** element contains the request to get a user's photo.</span></span> 
  
```XML
<GetUserPhoto>
   <Email/>
   <SizeRequested/>
</GetUserPhoto>
```

 <span data-ttu-id="c3a7b-105">**GetUserPhotoType**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-105">**GetUserPhotoType**</span></span>
## <a name="attributes-and-elements"></a><span data-ttu-id="c3a7b-106">属性および要素</span><span class="sxs-lookup"><span data-stu-id="c3a7b-106">Attributes and elements</span></span>

<span data-ttu-id="c3a7b-107">以下のセクションで、属性、子要素、親要素について説明します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-107">The following sections describe attributes, child elements, and parent elements.</span></span>
  
### <a name="attributes"></a><span data-ttu-id="c3a7b-108">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-108">Attributes</span></span>

<span data-ttu-id="c3a7b-109">なし。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-109">None.</span></span>
  
### <a name="child-elements"></a><span data-ttu-id="c3a7b-110">子要素</span><span class="sxs-lookup"><span data-stu-id="c3a7b-110">Child elements</span></span>

<span data-ttu-id="c3a7b-111">[電子メール (文字列)](email-string.md) | [SizeRequested](sizerequested.md)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-111">[Email (String)](email-string.md) | [SizeRequested](sizerequested.md)</span></span>
  
### <a name="parent-elements"></a><span data-ttu-id="c3a7b-112">親要素</span><span class="sxs-lookup"><span data-stu-id="c3a7b-112">Parent elements</span></span>

<span data-ttu-id="c3a7b-113">なし。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-113">None.</span></span>
  
## <a name="remarks"></a><span data-ttu-id="c3a7b-114">備考</span><span class="sxs-lookup"><span data-stu-id="c3a7b-114">Remarks</span></span>

<span data-ttu-id="c3a7b-115">この要素は Exchange Server 2013 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-115">This element was introduced in Exchange Server 2013.</span></span>
  
<span data-ttu-id="c3a7b-116">この要素を記述するスキーマは、Exchange Web サービスをホストする IIS 仮想ディレクトリに置かれています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-116">The schema that describes this element is located in the IIS virtual directory that hosts Exchange Web Services.</span></span>
  
## <a name="element-information"></a><span data-ttu-id="c3a7b-117">要素情報</span><span class="sxs-lookup"><span data-stu-id="c3a7b-117">Element information</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="c3a7b-118">名前空間</span><span class="sxs-lookup"><span data-stu-id="c3a7b-118">Namespace</span></span>  <br/> |http://schemas.microsoft.com/exchange/services/2006/messages  <br/> |
|<span data-ttu-id="c3a7b-119">スキーマ名</span><span class="sxs-lookup"><span data-stu-id="c3a7b-119">Schema name</span></span>  <br/> |<span data-ttu-id="c3a7b-120">メッセージ スキーマ</span><span class="sxs-lookup"><span data-stu-id="c3a7b-120">Messages schema</span></span>  <br/> |
|<span data-ttu-id="c3a7b-121">検証ファイル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-121">Validation file</span></span>  <br/> |<span data-ttu-id="c3a7b-122">messages.xsd</span><span class="sxs-lookup"><span data-stu-id="c3a7b-122">messages.xsd</span></span>  <br/> |
|<span data-ttu-id="c3a7b-123">空にすることができます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-123">Can be empty</span></span>  <br/> ||
   
