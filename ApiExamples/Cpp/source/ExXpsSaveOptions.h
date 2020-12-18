#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
#include <system/io/file_info.h>
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
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExXpsSaveOptions : public ApiExampleBase
{
public:
    void OptimizeOutput(bool optimizeOutput)
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to xps.
        auto doc = MakeObject<Document>(MyDir + u"Unoptimized document.docx");

        // When saving to .xps, we can use SaveOptions to optimize the output in some cases
        auto saveOptions = MakeObject<XpsSaveOptions>();
        saveOptions->set_OptimizeOutput(optimizeOutput);

        doc->Save(ArtifactsDir + u"XpsSaveOptions.OptimizeOutput.xps", saveOptions);
        //ExEnd

        // The input document had adjacent runs with the same formatting, which, if output optimization was enabled,
        // have been combined to save space
        auto outFileInfo = MakeObject<System::IO::FileInfo>(ArtifactsDir + u"XpsSaveOptions.OptimizeOutput.xps");

        if (optimizeOutput)
        {
            ASSERT_TRUE(outFileInfo->get_Length() < 51000);
        }
        else
        {
            ASSERT_TRUE(outFileInfo->get_Length() > 60000);
        }
    }
};

} // namespace ApiExamples
