#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Underline.h>
#include <Model/Text/ParagraphFormat.h>
#include <Model/Text/ParagraphAlignment.h>
#include <Model/Text/Font.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <drawing/color.h>

using namespace Aspose::Words;

void DocumentBuilderInsertParagraph()
{
    // ExStart:DocumentBuilderInsertParagraph
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Specify font formatting
    System::SharedPtr<Font> font = builder->get_Font();
    font->set_Size(16);
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_Blue());
    font->set_Name(u"Arial");
    font->set_Underline(Aspose::Words::Underline::Dash);
    
    // Specify paragraph formatting
    System::SharedPtr<ParagraphFormat> paragraphFormat = builder->get_ParagraphFormat();
    paragraphFormat->set_FirstLineIndent(8);
    paragraphFormat->set_Alignment(Aspose::Words::ParagraphAlignment::Justify);
    paragraphFormat->set_KeepTogether(true);
    
    builder->Writeln(u"A whole paragraph.");
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderInsertParagraph.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderInsertParagraph
    std::cout << "\nParagraph inserted successfully into the document using DocumentBuilder.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
