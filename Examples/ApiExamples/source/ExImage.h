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

/// <summary>
/// Mostly scenarios that deal with image shapes.
/// </summary>
class ExImage : public ApiExampleBase
{
    typedef ExImage ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void FromFile();
    void FromUrl();
    void FromStream();
    void CreateFloatingPageCenter();
    void CreateFloatingPositionSize();
    void InsertImageWithHyperlink();
    void CreateLinkedImage();
    void DeleteAllImages();
    void DeleteAllImagesPreOrder();
    void ScaleImage();
    void InsertWebpImage();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


