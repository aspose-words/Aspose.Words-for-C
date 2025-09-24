#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Loading/TxtTrailingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Loading/TxtLeadingSpacesOptions.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Loading;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExTxtLoadOptions : public ApiExampleBase
{
    typedef ExTxtLoadOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces);
    void TrailSpaces(Aspose::Words::Loading::TxtLeadingSpacesOptions txtLeadingSpacesOptions, Aspose::Words::Loading::TxtTrailingSpacesOptions txtTrailingSpacesOptions);
    void DetectDocumentDirection();
    void AutoNumberingDetection();
    void DetectHyperlinks();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


