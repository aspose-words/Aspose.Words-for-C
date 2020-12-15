#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/ComHelper.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <system/details/dispose_guard.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/test_tools/compare.h>
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

class ExComHelper : public ApiExampleBase
{
public:
    void ComHelper_()
    {
        //ExStart
        //ExFor:ComHelper
        //ExFor:ComHelper.#ctor
        //ExFor:ComHelper.Open(Stream)
        //ExFor:ComHelper.Open(String)
        //ExSummary:Shows how to open documents using the ComHelper class.
        // The ComHelper class allows us to load documents from within COM clients.
        auto comHelper = MakeObject<ComHelper>();

        // 1 -  Using a local system filename:
        SharedPtr<Document> doc = comHelper->Open(MyDir + u"Document.docx");

        ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", doc->GetText().Trim());

        // 2 -  From a stream:
        {
            auto stream = MakeObject<System::IO::FileStream>(MyDir + u"Document.docx", System::IO::FileMode::Open);
            doc = comHelper->Open(stream);

            ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", doc->GetText().Trim());
        }
        //ExEnd
    }
};

} // namespace ApiExamples
