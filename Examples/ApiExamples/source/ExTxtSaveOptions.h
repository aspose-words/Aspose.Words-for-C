#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Saving/TxtExportHeadersFootersMode.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Saving;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExTxtSaveOptions : public ApiExampleBase
{
    typedef ExTxtSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void PageBreaks(bool forcePageBreaks);
    void AddBidiMarks(bool addBidiMarks);
    void ExportHeadersFooters(Aspose::Words::Saving::TxtExportHeadersFootersMode txtExportHeadersFootersMode);
    void TxtListIndentation();
    void SimplifyListLabels(bool simplifyListLabels);
    void ParagraphBreak();
    void Encoding();
    void PreserveTableLayout(bool preserveTableLayout);
    void MaxCharactersPerLine();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


