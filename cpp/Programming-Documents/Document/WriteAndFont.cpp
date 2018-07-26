#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <Model/Text/Underline.h>
#include <Model/Text/Font.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <drawing/color.h>

using namespace Aspose::Words;

void WriteAndFont()
{
    // ExStart:WriteAndFont
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Specify font formatting before adding text.
    System::SharedPtr<Font> font = builder->get_Font();
    font->set_Size(16);
    font->set_Bold(true);
    font->set_Color(System::Drawing::Color::get_Blue());
    font->set_Name(u"Arial");
    font->set_Underline(Aspose::Words::Underline::Dash);
    
    builder->Write(u"Sample text.");
    dataDir = dataDir + GetOutputFilePath(u"WriteAndFont.doc");
    doc->Save(dataDir);
    // ExEnd:WriteAndFont
    std::cout << "\nFormatted text using DocumentBuilder inserted successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
