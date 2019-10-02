#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Underline.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void WriteAndFont()
{
    std::cout << "WriteAndFont example started." << std::endl;
    // ExStart:WriteAndFont
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Specify font formatting before adding text.
    System::SharedPtr<Font> font = builder->get_Font();
    font->set_Size(16);
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_Blue());
    font->set_Name(u"Arial");
    font->set_Underline(Underline::Dash);

    builder->Write(u"Sample text.");
    System::String outputPath = outputDataDir + u"WriteAndFont.doc";
    doc->Save(outputPath);
    // ExEnd:WriteAndFont
    std::cout << "Formatted text using DocumentBuilder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "WriteAndFont example started." << std::endl << std::endl;
}
