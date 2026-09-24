#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Saving/RtfSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options {

class WorkingWithRtfSaveOptions : public DocsExamplesBase
{
public:
    void SavingImagesAsWmf()
    {
        //ExStart:SavingImagesAsWmf
        //GistId:a20e6220d716ce2f51299b6087df12bb
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<RtfSaveOptions>();
        saveOptions->set_SaveImagesAsWmf(true);

        doc->Save(ArtifactsDir + u"WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
        //ExEnd:SavingImagesAsWmf
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
