#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>

using namespace Aspose::Words;

void InsertStyleSeparator()
{
    std::cout << "InsertStyleSeparator example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithStyles();
    // ExStart:ParagraphInsertStyleSeparator
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    System::SharedPtr<Style> paraStyle = builder->get_Document()->get_Styles()->Add(StyleType::Paragraph, u"MyParaStyle");
    paraStyle->get_Font()->set_Bold(false);
    paraStyle->get_Font()->set_Size(8);
    paraStyle->get_Font()->set_Name(u"Arial");
    // Append text with "Heading 1" style.
    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
    builder->Write(u"Heading 1");
    builder->InsertStyleSeparator();
    // Append text with another style.
    builder->get_ParagraphFormat()->set_StyleName(paraStyle->get_Name());
    builder->Write(u"This is text with some other formatting ");
    System::String outputPath = outputDataDir + u"InsertStyleSeparator.doc";
    // Save the document to disk.
    doc->Save(outputPath);
    // ExEnd:ParagraphInsertStyleSeparator
    std::cout << "Applied different paragraph styles to two different parts of a text line successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertStyleSeparator example finished." << std::endl << std::endl;
}