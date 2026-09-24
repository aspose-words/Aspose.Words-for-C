#pragma once

#ifdef ASPOSE_BARCODE_AVAILABLE

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>

#include "CustomBarcodeGenerator.h"
#include "DocsExamplesBase.h"

using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements {

class WorkingWithBarcodeGenerator : public DocsExamplesBase
{
public:
    void BarcodeGenerator()
    {
        //ExStart:BarcodeGenerator
        //GistId:c9ce0847b39f05c1bc89daf438acc0bf
        auto doc = MakeObject<Document>(MyDir + u"Field sample - BARCODE.docx");

        doc->get_FieldOptions()->set_BarcodeGenerator(MakeObject<DocsExamples::CustomBarcodeGenerator>());

        doc->Save(ArtifactsDir + u"WorkingWithBarcodeGenerator.BarcodeGenerator.pdf");
        //ExEnd:BarcodeGenerator
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements

#endif // ASPOSE_BARCODE_AVAILABLE
