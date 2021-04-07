#pragma once

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment {

class WorkingWithRanges : public DocsExamplesBase
{
public:
    void RangesDeleteText()
    {
        //ExStart:RangesDeleteText
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        doc->get_Sections()->idx_get(0)->get_Range()->Delete();
        //ExEnd:RangesDeleteText
    }

    void RangesGetText()
    {
        //ExStart:RangesGetText
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        String text = doc->get_Range()->get_Text();
        //ExEnd:RangesGetText
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment
