#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/string.h>

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace ApiExamples {

class ExMossRtf2Docx
{
public:
    static void ConvertRtfToDocx(String inFileName, String outFileName)
    {
        // Load an RTF file into Aspose.Words.
        auto doc = MakeObject<Document>(inFileName);

        // Save the document in the OOXML format.
        doc->Save(outFileName, SaveFormat::Docx);
    }
    ExMossRtf2Docx() = delete;
};

} // namespace ApiExamples
