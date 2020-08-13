// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMossRtf2Docx.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

void ExMossRtf2Docx::ConvertRtfToDocx(String inFileName, String outFileName)
{
    // Load an RTF file into Aspose.Words.
    auto doc = MakeObject<Document>(inFileName);

    // Save the document in the OOXML format.
    doc->Save(outFileName, Aspose::Words::SaveFormat::Docx);
}

} // namespace ApiExamples
