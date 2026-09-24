#pragma once

#include <gtest/gtest.h>
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
    void ExportOfficeMathAsLatex();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


