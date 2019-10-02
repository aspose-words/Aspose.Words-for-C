#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMathDisplayType.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMathJustification.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Math;

void UseOfficeMathProperties()
{
    std::cout << "UseOfficeMathProperties example started." << std::endl;
    // ExStart:UseOfficeMathProperties
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"MathEquations.docx");
    System::SharedPtr<OfficeMath> officeMath = System::DynamicCast<OfficeMath>(doc->GetChild(NodeType::OfficeMath, 0, true));

    // Gets/sets Office Math display format type which represents whether an equation is displayed inline with the text or displayed on its own line.
    officeMath->set_DisplayType(OfficeMathDisplayType::Display);
    // or OfficeMathDisplayType.Inline

    // Gets/sets Office Math justification.
    officeMath->set_Justification(OfficeMathJustification::Left);
    // Left justification of Math Paragraph.

    System::String outputPath = outputDataDir + u"UseOfficeMathProperties.docx";
    doc->Save(outputPath);
    // ExEnd:UseOfficeMathProperties
    std::cout << "UseOfficeMathProperties example finished." << std::endl << std::endl;
}