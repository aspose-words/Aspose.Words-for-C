#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Saving { enum class TxtExportHeadersFootersMode; } } }

namespace ApiExamples {

class ExTxtSaveOptions : public ApiExampleBase
{
public:

    void PageBreaks(bool forcePageBreaks);
    void AddBidiMarks(bool addBidiMarks);
    void ExportHeadersFooters(Aspose::Words::Saving::TxtExportHeadersFootersMode txtExportHeadersFootersMode);
    void TxtListIndentation();
    void SimplifyListLabels(bool simplifyListLabels);
    void ParagraphBreak();
    void Encoding();
    void TableLayout(bool preserveTableLayout);
    
};

} // namespace ApiExamples


