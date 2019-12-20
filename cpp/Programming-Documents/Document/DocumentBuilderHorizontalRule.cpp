#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleAlignment.h>
#include <Aspose.Words.Cpp/Model/Drawing/HorizontalRuleFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{
    void DocumentBuilderInsertHorizontalRule(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderInsertHorizontalRule
        // Initialize document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Insert a horizontal rule shape into the document.");
        builder->InsertHorizontalRule();

        System::String outputPath = outputDataDir + u"DocumentBuilderHorizontalRule.DocumentBuilderInsertHorizontalRule.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderInsertHorizontalRule
        std::cout << "Horizontal rule is inserted into document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void DocumentBuilderHorizontalRuleFormat(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderHorizontalRuleFormat
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>();

        System::SharedPtr<Shape> shape = builder->InsertHorizontalRule();
        System::SharedPtr<HorizontalRuleFormat> horizontalRuleFormat = shape->get_HorizontalRuleFormat();

        horizontalRuleFormat->set_Alignment(HorizontalRuleAlignment::Center);
        horizontalRuleFormat->set_WidthPercent(70);
        horizontalRuleFormat->set_Height(3);
        horizontalRuleFormat->set_Color(System::Drawing::Color::get_Blue());
        horizontalRuleFormat->set_NoShade(true);

        System::String outputPath = outputDataDir + u"DocumentBuilderHorizontalRule.DocumentBuilderHorizontalRuleFormat.doc";
        builder->get_Document()->Save(outputPath);
        // ExEnd:DocumentBuilderHorizontalRuleFormat
        std::cout << "Horizontal rule format inserted into document successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void DocumentBuilderHorizontalRule()
{
    std::cout << "DocumentBuilderHorizontalRule example started." << std::endl;
    // The path to the output documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

    DocumentBuilderInsertHorizontalRule(outputDataDir);
    DocumentBuilderHorizontalRuleFormat(outputDataDir);
    std::cout << "DocumentBuilderHorizontalRule example finished." << std::endl << std::endl;
}