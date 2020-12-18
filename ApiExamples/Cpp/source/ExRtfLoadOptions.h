#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/RtfLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;

namespace ApiExamples {

class ExRtfLoadOptions : public ApiExampleBase
{
public:
    void RecognizeUtf8Text(bool doRecognizeUtb8Text)
    {
        //ExStart
        //ExFor:RtfLoadOptions
        //ExFor:RtfLoadOptions.#ctor
        //ExFor:RtfLoadOptions.RecognizeUtf8Text
        //ExSummary:Shows how to detect UTF8 characters during import.
        auto loadOptions = MakeObject<RtfLoadOptions>();
        loadOptions->set_RecognizeUtf8Text(doRecognizeUtb8Text);

        auto doc = MakeObject<Document>(MyDir + u"UTF-8 characters.rtf", loadOptions);

        ASSERT_EQ(doRecognizeUtb8Text ? String(u"“John Doe´s list of currency symbols”™\r€, ¢, £, ¥, ¤")
                                      : String(u"â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\râ‚¬, Â¢, Â£, Â¥, Â¤"),
                  doc->get_FirstSection()->get_Body()->GetText().Trim());
        //ExEnd
    }
};

} // namespace ApiExamples
