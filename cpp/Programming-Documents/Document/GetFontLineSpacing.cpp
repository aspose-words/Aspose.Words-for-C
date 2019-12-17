#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Text/Font.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/Run.h>
#include <Model/Text/RunCollection.h>

using namespace Aspose::Words;

void GetFontLineSpacing()
{
    std::cout << "GetFontLineSpacing example started." << std::endl;
    // ExStart:GetFontLineSpacing
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->get_Font()->set_Name(u"Calibri");
    builder->Writeln(u"qText");

    // Obtain line spacing.
    System::SharedPtr<Font> font = builder->get_Document()->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font();
    //Console.WriteLine($"lineSpacing = {font.LineSpacing}");
    std::cout << "lineSpacing = " << font->get_LineSpacing() << std::endl;
    // ExEnd:GetFontLineSpacing
    std::cout << "GetFontLineSpacing example finished." << std::endl << std::endl;
}