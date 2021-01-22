#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <drawing/color.h>

using namespace Aspose::Words;

void DocumentBuilderInsertParagraph()
{
    std::cout << "DocumentBuilderInsertParagraph example started." << std::endl;
    // ExStart:DocumentBuilderInsertParagraph
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Specify font formatting
    System::SharedPtr<Font> font = builder->get_Font();
    font->set_Size(16);
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_Blue());
    font->set_Name(u"Arial");
    font->set_Underline(Underline::Dash);

    // Specify paragraph formatting
    System::SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
    paragraphFormat->set_FirstLineIndent(8);
    paragraphFormat->set_Alignment(ParagraphAlignment::Justify);
    paragraphFormat->set_KeepTogether(true);

    builder->Writeln(u"A whole paragraph.");
    System::String outputPath = outputDataDir + u"DocumentBuilderInsertParagraph.doc";
    doc->Save(outputPath);
    // ExEnd:DocumentBuilderInsertParagraph
    std::cout << "Paragraph inserted successfully into the document using DocumentBuilder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "DocumentBuilderInsertParagraph example finished." << std::endl << std::endl;
}
