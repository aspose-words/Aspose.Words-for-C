#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { enum class TxtLeadingSpacesOptions; } }
namespace Aspose { namespace Words { enum class TxtTrailingSpacesOptions; } }

namespace ApiExamples {

class ExTxtLoadOptions : public ApiExampleBase
{
public:

    void DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces);
    void TrailSpaces(Aspose::Words::TxtLeadingSpacesOptions txtLeadingSpacesOptions, Aspose::Words::TxtTrailingSpacesOptions txtTrailingSpacesOptions);
    void DetectDocumentDirection();
    
};

} // namespace ApiExamples


