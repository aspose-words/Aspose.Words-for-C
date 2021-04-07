// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "DocumentHelper.h"

#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Tables/AutoFitBehavior.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <system/details/dispose_guard.h>
#include <system/io/memory_stream.h>
#include <system/io/stream_reader.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <gtest/gtest.h>
#include <testing/test_predicates.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

System::SharedPtr<Document> DocumentHelper::CreateDocumentWithoutDummyText()
{
    auto doc = System::MakeObject<Document>();

    // Remove the previous changes of the document
    doc->RemoveAllChildren();

    // Set the document author
    doc->get_BuiltInDocumentProperties()->set_Author(u"Test Author");

    // Create paragraph without run
    auto builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Writeln();

    return doc;
}

System::SharedPtr<Document> DocumentHelper::CreateDocumentFillWithDummyText()
{
    auto doc = System::MakeObject<Document>();

    // Remove the previous changes of the document
    doc->RemoveAllChildren();

    // Set the document author
    doc->get_BuiltInDocumentProperties()->set_Author(u"Test Author");

    auto builder = System::MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Page ");
    builder->InsertField(u"PAGE", u"");
    builder->Write(u" of ");
    builder->InsertField(u"NUMPAGES", u"");

    // Insert new table with two rows and two cells
    InsertTable(builder);

    builder->Writeln(u"Hello World!");

    // Continued on page 2 of the document content
    builder->InsertBreak(BreakType::PageBreak);

    // Insert TOC entries
    InsertToc(builder);

    return doc;
}

void DocumentHelper::FindTextInFile(System::String path, System::String expression)
{
    {
        auto sr = System::MakeObject<System::IO::StreamReader>(path);
        while (!sr->get_EndOfStream())
        {
            System::String line = sr->ReadLine();

            if (System::String::IsNullOrEmpty(line))
            {
                continue;
            }

            if (line.Contains(expression))
            {
                std::cout << line << std::endl;
                SUCCEED();
                return;
            }
            else
            {
                FAIL();
            }
        }
    }
}

System::SharedPtr<Document> DocumentHelper::CreateSimpleDocument(System::String templateText)
{
    auto doc = System::MakeObject<Document>();
    auto builder = System::MakeObject<DocumentBuilder>(doc);

    builder->Write(templateText);

    return doc;
}

System::SharedPtr<Document> DocumentHelper::CreateTemplateDocumentWithDrawObjects(System::String templateText, ShapeType shapeType)
{
    auto doc = System::MakeObject<Document>();

    // Create textbox shape.
    auto shape = System::MakeObject<Shape>(doc, shapeType);
    shape->set_Width(431.5);
    shape->set_Height(346.35);

    auto paragraph = System::MakeObject<Paragraph>(doc);
    paragraph->AppendChild(System::MakeObject<Run>(doc, templateText));

    // Insert paragraph into the textbox.
    shape->AppendChild(paragraph);

    // Insert textbox into the document.
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(shape);

    return doc;
}

bool DocumentHelper::CompareDocs(System::String filePathDoc1, System::String filePathDoc2)
{
    auto doc1 = System::MakeObject<Document>(filePathDoc1);
    auto doc2 = System::MakeObject<Document>(filePathDoc2);

    if (doc1->GetText() == doc2->GetText())
    {
        return true;
    }

    return false;
}

System::SharedPtr<Run> DocumentHelper::InsertNewRun(System::SharedPtr<Document> doc, System::String text, int32_t paraIndex)
{
    System::SharedPtr<Paragraph> para = GetParagraph(doc, paraIndex);

    auto run = System::MakeObject<Run>(doc);
    run->set_Text(text);

    para->AppendChild(run);

    return run;
}

void DocumentHelper::InsertBuilderText(System::SharedPtr<DocumentBuilder> builder, System::ArrayPtr<System::String> textStrings)
{
    for (System::String textString : textStrings)
    {
        builder->Writeln(textString);
    }
}

System::String DocumentHelper::GetParagraphText(System::SharedPtr<Document> doc, int32_t paraIndex)
{
    return doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(paraIndex)->GetText();
}

System::SharedPtr<Table> DocumentHelper::InsertTable(System::SharedPtr<DocumentBuilder> builder)
{
    // Start creating a new table
    System::SharedPtr<Table> table = builder->StartTable();

    // Insert Row 1 Cell 1
    builder->InsertCell();
    builder->Write(u"Date");

    // Set width to fit the table contents
    table->AutoFit(AutoFitBehavior::AutoFitToContents);

    // Insert Row 1 Cell 2
    builder->InsertCell();
    builder->Write(u" ");

    builder->EndRow();

    // Insert Row 2 Cell 1
    builder->InsertCell();
    builder->Write(u"Author");

    // Insert Row 2 Cell 2
    builder->InsertCell();
    builder->Write(u" ");

    builder->EndRow();

    builder->EndTable();

    return table;
}

void DocumentHelper::InsertToc(System::SharedPtr<DocumentBuilder> builder)
{
    // Creating TOC entries
    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

    builder->Writeln(u"Heading 1");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

    builder->Writeln(u"Heading 1.1");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading4);

    builder->Writeln(u"Heading 1.1.1.1");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading5);

    builder->Writeln(u"Heading 1.1.1.1.1");

    builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading9);

    builder->Writeln(u"Heading 1.1.1.1.1.1.1.1.1");
}

System::String DocumentHelper::GetSectionText(System::SharedPtr<Document> doc, int32_t secIndex)
{
    return doc->get_Sections()->idx_get(secIndex)->GetText();
}

System::SharedPtr<Paragraph> DocumentHelper::GetParagraph(System::SharedPtr<Document> doc, int32_t paraIndex)
{
    return doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(paraIndex);
}

System::SharedPtr<Document> DocumentHelper::SaveOpen(System::SharedPtr<Document> doc)
{
    {
        auto docStream = System::MakeObject<System::IO::MemoryStream>();
        doc->Save(docStream, System::MakeObject<OoxmlSaveOptions>(SaveFormat::Docx));
        return System::MakeObject<Document>(docStream);
    }
}

} // namespace ApiExamples
