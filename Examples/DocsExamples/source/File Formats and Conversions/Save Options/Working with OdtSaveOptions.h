#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Saving/OdtSaveMeasureUnit.h>
#include <Aspose.Words.Cpp/Saving/OdtSaveOptions.h>
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

class WorkingWithOdtSaveOptions : public DocsExamplesBase
{
public:
    void MeasureUnit()
    {
        //ExStart:MeasureUnit
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
        // and content properties in documents whereas MS Office uses inches.
        auto saveOptions = MakeObject<OdtSaveOptions>();
        saveOptions->set_MeasureUnit(OdtSaveMeasureUnit::Inches);

        doc->Save(ArtifactsDir + u"WorkingWithOdtSaveOptions.MeasureUnit.odt", saveOptions);
        //ExEnd:MeasureUnit
    }
};

}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options
