#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Loading/RtfLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;

namespace ApiExamples {

class ExRtfLoadOptions : public ApiExampleBase
{
public:
    void RecognizeUtf8Text(bool recognizeUtf8Text)
    {
        //ExStart
        //ExFor:RtfLoadOptions
        //ExFor:RtfLoadOptions.#ctor
        //ExFor:RtfLoadOptions.RecognizeUtf8Text
        //ExSummary:Shows how to detect UTF-8 characters while loading an RTF document.
        // Create an "RtfLoadOptions" object to modify how we load an RTF document.
        auto loadOptions = MakeObject<RtfLoadOptions>();

        // Set the "RecognizeUtf8Text" property to "false" to assume that the document uses the ISO 8859-1 charset
        // and loads every character in the document.
        // Set the "RecognizeUtf8Text" property to "true" to parse any variable-length characters that may occur in the text.
        loadOptions->set_RecognizeUtf8Text(recognizeUtf8Text);

        auto doc = MakeObject<Document>(MyDir + u"UTF-8 characters.rtf", loadOptions);

        ASSERT_EQ(recognizeUtf8Text ? String(u"“John Doe´s list of currency symbols”™\r") + u"€, ¢, £, ¥, ¤"
                                    : String(u"â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\r") + u"â‚¬, Â¢, Â£, Â¥, Â¤",
                  doc->get_FirstSection()->get_Body()->GetText().Trim());
        //ExEnd
    }
};

} // namespace ApiExamples
