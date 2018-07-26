#include "../../examples.h"

#include <system/shared_ptr.h>

#include "Model/Document/Document.h"
#include "Model/Document/DocumentBuilder.h"
#include <Model/Styles/Style.h>
#include <Model/Styles/StyleIdentifier.h>
#include <Model/Styles/StyleCollection.h>
#include <Model/Styles/StyleType.h>
#include <Model/Text/Font.h>
#include <Model/Text/ParagraphFormat.h>

using namespace Aspose::Words;

void InsertStyleSeparator()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithStyles();
    // ExStart:ParagraphInsertStyleSeparator
    auto doc = System::MakeObject<Document>();
    auto builder = System::MakeObject<DocumentBuilder>(doc);
    auto paraStyle = builder->get_Document()->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyParaStyle");
    paraStyle->get_Font()->set_Bold(false);
    paraStyle->get_Font()->set_Size(8);
    paraStyle->get_Font()->set_Name(u"Arial");
    // Append text with "Heading 1" style.
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Write(u"Heading 1");
    builder->InsertStyleSeparator();
    // Append text with another style.
    builder->get_ParagraphFormat()->set_StyleName(paraStyle->get_Name());
    builder->Write(u"This is text with some other formatting ");
    System::String outputPath = dataDir + GetOutputFilePath(u"InsertStyleSeparator.doc");
    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:ParagraphInsertStyleSeparator
    std::cout << "\nApplied different paragraph styles to two different parts of a text line successfully.\nFile saved at " << outputPath.ToUtf8String() << std::endl;
}