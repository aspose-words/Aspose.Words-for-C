#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/PclSaveOptions.h>
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

class WorkingWithPclSaveOptions : public DocsExamplesBase
{
public:
    void RasterizeTransformedElements()
    {
        //ExStart:RasterizeTransformedElements
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PclSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Pcl);
        saveOptions->set_RasterizeTransformedElements(false);

        doc->Save(ArtifactsDir + u"WorkingWithPclSaveOptions.RasterizeTransformedElements.pcl", saveOptions);
        //ExEnd:RasterizeTransformedElements
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
