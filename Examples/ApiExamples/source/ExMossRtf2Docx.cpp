// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMossRtf2Docx.h"

#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

void ExMossRtf2Docx::ConvertRtfToDocx(System::String inFileName, System::String outFileName)
{
    // Load an RTF file into Aspose.Words.
    auto doc = System::MakeObject<Aspose::Words::Document>(inFileName);
    
    // Save the document in the OOXML format.
    doc->Save(outFileName, Aspose::Words::SaveFormat::Docx);
}

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
