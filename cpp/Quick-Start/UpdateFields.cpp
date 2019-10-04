#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>

using namespace Aspose::Words;

void UpdateFields()
{
    std::cout << "UpdateFields example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_QuickStart();

    // Demonstrates how to insert fields and update them using Aspose.Words.

    // First create a blank document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    // Use the document builder to insert some content and fields.
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert a table of contents at the beginning of the document.
    builder->InsertTableOfContents(u"\\o \"1-3\" \\h \\z \\u");
    builder->Writeln();

    // Insert some other fields.
    builder->Write(u"Page: ");
    builder->InsertField(u"PAGE");
    builder->Write(u" of ");
    builder->InsertField(u"NUMPAGES");
    builder->Writeln();

    builder->Write(u"Date: ");
    builder->InsertField(u"DATE");

    // Start the actual document content on the second page.
    builder->InsertBreak(BreakType::SectionBreakNewPage);

    // Build a document with complex structure by applying different heading styles thus creating TOC entries.
    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

    builder->Writeln(u"Heading 1");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

    builder->Writeln(u"Heading 1.1");
    builder->Writeln(u"Heading 1.2");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

    builder->Writeln(u"Heading 2");
    builder->Writeln(u"Heading 3");

    // Move to the next page.
    builder->InsertBreak(BreakType::PageBreak);

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

    builder->Writeln(u"Heading 3.1");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading3);

    builder->Writeln(u"Heading 3.1.1");
    builder->Writeln(u"Heading 3.1.2");
    builder->Writeln(u"Heading 3.1.3");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

    builder->Writeln(u"Heading 3.2");
    builder->Writeln(u"Heading 3.3");

    std::cout << "Updating all fields in the document." << std::endl;

    // Call the method below to update the TOC.
    doc->UpdateFields();

    System::String outputPath = outputDataDir + u"UpdateFields.docx";
    doc->Save(outputPath);

    std::cout << "Fields updated successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "UpdateFields example finished." << std::endl << std::endl;
}