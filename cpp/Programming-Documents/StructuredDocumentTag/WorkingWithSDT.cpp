#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtType.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/XmlMapping.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPart.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPartCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/MarkupLevel.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagRangeStart.h>
#include <system/enumerator_adapter.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Tables;

namespace
{
    void SetContentControlColor(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetContentControlColor

        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->set_Color(System::Drawing::Color::get_Red());

        System::String outputPath = outputDataDir + u"WorkingWithSDT.SetContentControlColor.docx";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetContentControlColor
        std::cout << "Set the color of content control successfully." << std::endl;
    }

    void ClearContentsControl(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:ClearContentsControl
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->Clear();

        System::String outputPath = outputDataDir + u"WorkingWithSDT.ClearContentsControl.doc";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:ClearContentsControl
        std::cout << "Clear the contents of content control successfully." << std::endl;
    }

    void BindSDTtoCustomXmlPart(System::String const &outputDataDir)
    {
        // ExStart:BindSDTtoCustomXmlPart
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(System::Guid::NewGuid().ToString(u"B"), u"<root><text>Hello, World!</text></root>");

        System::SharedPtr<StructuredDocumentTag> sdt = System::MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(sdt);
        sdt->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[1]", u"");

        System::String outputPath = outputDataDir + u"WorkingWithSDT.BindSDTtoCustomXmlPart.doc";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:BindSDTtoCustomXmlPart
        std::cout << "Creation of an XML part and binding a content control to it successfully." << std::endl;
    }

    void SetContentControlStyle(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetContentControlStyle
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"input.docx");
        System::SharedPtr<StructuredDocumentTag> sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        System::SharedPtr<Style> style = doc->get_Styles()->idx_get(StyleIdentifier::Quote);
        sdt->set_Style(style);

        System::String outputPath = outputDataDir + u"WorkingWithSDT.SetContentControlStyle.docx";
        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:SetContentControlStyle
        std::cout << "Set the style of content control successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void CreatingTableRepeatingSectionMappedToCustomXmlPart(System::String const &outputDataDir)
    {
        // ExStart:CreatingTableRepeatingSectionMappedToCustomXmlPart
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::String xml = System::String(u"<books><book><title>Everyday Italian</title><author>Giada De Laurentiis</author></book>") +
                             u"<book><title>Harry Potter</title><author>J K. Rowling</author></book>" +
                             u"<book><title>Learning XML</title><author>Erik T. Ray</author></book></books>";
        System::SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(u"Books", xml);

        System::SharedPtr<Table> table = builder->StartTable();

        builder->InsertCell();
        builder->Write(u"Title");

        builder->InsertCell();
        builder->Write(u"Author");

        builder->EndRow();
        builder->EndTable();

        System::SharedPtr<StructuredDocumentTag> repeatingSectionSdt = System::MakeObject<StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::RepeatingSection, Aspose::Words::Markup::MarkupLevel::Row);
        repeatingSectionSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book", u"");
        table->AppendChild(repeatingSectionSdt);

        System::SharedPtr<StructuredDocumentTag> repeatingSectionItemSdt = System::MakeObject<StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::RepeatingSectionItem, Aspose::Words::Markup::MarkupLevel::Row);
        repeatingSectionSdt->AppendChild(repeatingSectionItemSdt);

        System::SharedPtr<Row> row = System::MakeObject<Row>(doc);
        repeatingSectionItemSdt->AppendChild(row);

        System::SharedPtr<StructuredDocumentTag> titleSdt = System::MakeObject<StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Cell);
        titleSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/title[1]", u"");
        row->AppendChild(titleSdt);

        System::SharedPtr<StructuredDocumentTag> authorSdt = System::MakeObject<StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Cell);
        authorSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/author[1]", u"");
        row->AppendChild(authorSdt);

        doc->Save(outputDataDir + u"WorkingWithSDT.CreatingTableRepeatingSectionMappedToCustomXmlPart.docx");
        // ExEnd:CreatingTableRepeatingSectionMappedToCustomXmlPart
        System::Console::WriteLine(u"\nCreation of a Table Repeating Section Mapped To a Custom Xml Part is successfull.");
    }


	void MultiSectionSDT(System::String const &inputDataDir, System::String const &outputDataDir)
	{
		// ExStart:MultiSectionSDT
		auto doc = System::MakeObject<Document>(inputDataDir + u"input.docx");
        auto tags = doc->GetChildNodes(NodeType::StructuredDocumentTagRangeStart, true);

		for (auto tag : System::IterateOver<StructuredDocumentTagRangeStart>(tags))
		{
            std::cout << tag->get_Title().ToUtf8String() << '\n';
		}
		// ExEnd:MultiSectionSDT
	}
}

void WorkingWithSDT()
{
    std::cout << "WorkingWithSDT example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithStructuredDocumentTag();
    System::String outputDataDir = GetOutputDataDir_WorkingWithStructuredDocumentTag();
    BindSDTtoCustomXmlPart(outputDataDir);
    ClearContentsControl(inputDataDir, outputDataDir);
    SetContentControlColor(inputDataDir, outputDataDir);
    SetContentControlStyle(inputDataDir, outputDataDir);
    CreatingTableRepeatingSectionMappedToCustomXmlPart(outputDataDir);
	MultiSectionSDT(inputDataDir, outputDataDir);
    std::cout << "WorkingWithSDT example finished." << std::endl << std::endl;
}