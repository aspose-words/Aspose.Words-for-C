#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingMode.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExImageSaveOptions : public ApiExampleBase
{
    typedef ExImageSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void OnePage();
    void Renderer(bool useGdiEmfRenderer);
    void PageSet();
    void WindowsMetaFile(Aspose::Words::Saving::MetafileRenderingMode metafileRenderingMode);
    void PageByPage();
    void PaperColor();
    void FloydSteinbergDithering();
    void EditImage();
    void JpegQuality();
    void Resolution();
    void ExportVariousPageRanges();
    void RenderInkObject();
    void GridLayout();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


