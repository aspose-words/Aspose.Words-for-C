#pragma once

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/RtfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>

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
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto saveOptions = MakeObject<RtfSaveOptions>();
        saveOptions->set_SaveImagesAsWmf(true);

        doc->Save(ArtifactsDir + u"WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
        //ExEnd:SavingImagesAsWmf
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
