#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Math/OfficeMathDisplayType.h>
#include <Aspose.Words.Cpp/Math/OfficeMathJustification.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Math;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements {

class WorkingWithOfficeMath : public DocsExamplesBase
{
public:
    void MathEquations()
    {
        //ExStart:MathEquations
        auto doc = MakeObject<Document>(MyDir + u"Office math.docx");
        auto officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));

        // OfficeMath display type represents whether an equation is displayed inline with the text or displayed on its line.
        officeMath->set_DisplayType(OfficeMathDisplayType::Display);
        officeMath->set_Justification(OfficeMathJustification::Left);

        doc->Save(ArtifactsDir + u"WorkingWithOfficeMath.MathEquations.docx");
        //ExEnd:MathEquations
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements
