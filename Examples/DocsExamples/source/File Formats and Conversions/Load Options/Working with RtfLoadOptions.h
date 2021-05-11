#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Loading/RtfLoadOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;

namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Load_Options {

class WorkingWithRtfLoadOptions : public DocsExamplesBase
{
public:
    void RecognizeUtf8Text()
    {
        //ExStart:RecognizeUtf8Text
        auto loadOptions = MakeObject<RtfLoadOptions>();
        loadOptions->set_RecognizeUtf8Text(true);

        auto doc = MakeObject<Document>(MyDir + u"UTF-8 characters.rtf", loadOptions);

        doc->Save(ArtifactsDir + u"WorkingWithRtfLoadOptions.RecognizeUtf8Text.rtf");
        //ExEnd:RecognizeUtf8Text
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Load_Options
