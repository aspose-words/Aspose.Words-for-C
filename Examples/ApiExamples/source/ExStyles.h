#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExStyles : public ApiExampleBase
{
    typedef ExStyles ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void Styles();
    void CreateStyle();
    void StyleCollection();
    void RemoveStylesFromStyleGallery();
    void ChangeTocsTabStops();
    void CopyStyleSameDocument();
    void CopyStyleDifferentDocument();
    void DefaultStyles();
    void ParagraphStyleBulletedList();
    void StyleAliases();
    void LockStyle();
    void StylePriority();
    void LinkedStyleName();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


