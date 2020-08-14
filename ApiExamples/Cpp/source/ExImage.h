#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include "ApiExampleBase.h"

namespace ApiExamples {

/// <summary>
/// Mostly scenarios that deal with image shapes.
/// </summary>
class ExImage : public ApiExampleBase
{
public:

    void CreateImageDirectly();
    void CreateFromUrl();
    void CreateFromStream();
    void CreateFromImage();
    void CreateFloatingPageCenter();
    void CreateFloatingPositionSize();
    void InsertImageWithHyperlink();
    void CreateLinkedImage();
    void DeleteAllImages();
    void DeleteAllImagesPreOrder();
    void ScaleImage();
    
};

} // namespace ApiExamples


