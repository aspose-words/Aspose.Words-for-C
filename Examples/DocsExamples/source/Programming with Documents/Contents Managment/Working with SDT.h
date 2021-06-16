#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Markup/CustomXmlPart.h>
#include <Aspose.Words.Cpp/Markup/CustomXmlPartCollection.h>
#include <Aspose.Words.Cpp/Markup/MarkupLevel.h>
#include <Aspose.Words.Cpp/Markup/SdtListItem.h>
#include <Aspose.Words.Cpp/Markup/SdtListItemCollection.h>
#include <Aspose.Words.Cpp/Markup/SdtType.h>
#include <Aspose.Words.Cpp/Markup/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Markup/StructuredDocumentTagRangeStart.h>
#include <Aspose.Words.Cpp/Markup/XmlMapping.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <drawing/color.h>
#include <system/array.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/guid.h>
#include <system/text/encoding.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment {

class WorkingWithSdt : public DocsExamplesBase
{
public:
    void CheckBoxTypeContentControl()
    {
        //ExStart:CheckBoxTypeContentControl
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto sdtCheckBox = MakeObject<StructuredDocumentTag>(doc, SdtType::Checkbox, MarkupLevel::Inline);
        builder->InsertNode(sdtCheckBox);

        doc->Save(ArtifactsDir + u"WorkingWithSdt.CheckBoxTypeContentControl.docx", SaveFormat::Docx);
        //ExEnd:CheckBoxTypeContentControl
    }

    void CurrentStateOfCheckBox()
    {
        //ExStart:SetCurrentStateOfCheckBox
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");

        // Get the first content control from the document.
        auto sdtCheckBox = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));

        if (sdtCheckBox->get_SdtType() == SdtType::Checkbox)
        {
            sdtCheckBox->set_Checked(true);
        }

        doc->Save(ArtifactsDir + u"WorkingWithSdt.CurrentStateOfCheckBox.docx");
        //ExEnd:SetCurrentStateOfCheckBox
    }

    void ModifyContentControls()
    {
        //ExStart:ModifyContentControls
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");

        for (const auto& sdt : System::IterateOver<StructuredDocumentTag>(doc->GetChildNodes(NodeType::StructuredDocumentTag, true)))
        {
            switch (sdt->get_SdtType())
            {
            case SdtType::PlainText: {
                sdt->RemoveAllChildren();
                auto para = System::DynamicCast_noexcept<Paragraph>(sdt->AppendChild(MakeObject<Paragraph>(doc)));
                auto run = MakeObject<Run>(doc, u"new text goes here");
                para->AppendChild(run);
                break;
            }

            case SdtType::DropDownList: {
                SharedPtr<SdtListItem> secondItem = sdt->get_ListItems()->idx_get(2);
                sdt->get_ListItems()->set_SelectedValue(secondItem);
                break;
            }

            case SdtType::Picture: {
                auto shape = System::DynamicCast<Shape>(sdt->GetChild(NodeType::Shape, 0, true));
                if (shape->get_HasImage())
                {
                    shape->get_ImageData()->SetImage(ImagesDir + u"Watermark.png");
                }

                break;
            }

            default:
                break;
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithSdt.ModifyContentControls.docx");
        //ExEnd:ModifyContentControls
    }

    void ComboBoxContentControl()
    {
        //ExStart:ComboBoxContentControl
        auto doc = MakeObject<Document>();

        auto sdt = MakeObject<StructuredDocumentTag>(doc, SdtType::ComboBox, MarkupLevel::Block);
        sdt->get_ListItems()->Add(MakeObject<SdtListItem>(u"Choose an item", u"-1"));
        sdt->get_ListItems()->Add(MakeObject<SdtListItem>(u"Item 1", u"1"));
        sdt->get_ListItems()->Add(MakeObject<SdtListItem>(u"Item 2", u"2"));
        doc->get_FirstSection()->get_Body()->AppendChild(sdt);

        doc->Save(ArtifactsDir + u"WorkingWithSdt.ComboBoxContentControl.docx");
        //ExEnd:ComboBoxContentControl
    }

    void RichTextBoxContentControl()
    {
        //ExStart:RichTextBoxContentControl
        auto doc = MakeObject<Document>();

        auto sdtRichText = MakeObject<StructuredDocumentTag>(doc, SdtType::RichText, MarkupLevel::Block);

        auto para = MakeObject<Paragraph>(doc);
        auto run = MakeObject<Run>(doc);
        run->set_Text(u"Hello World");
        run->get_Font()->set_Color(System::Drawing::Color::get_Green());
        para->get_Runs()->Add(run);
        sdtRichText->get_ChildNodes()->Add(para);
        doc->get_FirstSection()->get_Body()->AppendChild(sdtRichText);

        doc->Save(ArtifactsDir + u"WorkingWithSdt.RichTextBoxContentControl.docx");
        //ExEnd:RichTextBoxContentControl
    }

    void SetContentControlColor()
    {
        //ExStart:SetContentControlColor
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");

        auto sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->set_Color(System::Drawing::Color::get_Red());

        doc->Save(ArtifactsDir + u"WorkingWithSdt.SetContentControlColor.docx");
        //ExEnd:SetContentControlColor
    }

    void ClearContentsControl()
    {
        //ExStart:ClearContentsControl
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");

        auto sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        sdt->Clear();

        doc->Save(ArtifactsDir + u"WorkingWithSdt.ClearContentsControl.doc");
        //ExEnd:ClearContentsControl
    }

    void BindSdTtoCustomXmlPart()
    {
        //ExStart:BindSDTtoCustomXmlPart
        auto doc = MakeObject<Document>();
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(System::Guid::NewGuid().ToString(u"B"), u"<root><text>Hello, World!</text></root>");

        auto sdt = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(sdt);

        sdt->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[1]", u"");

        doc->Save(ArtifactsDir + u"WorkingWithSdt.BindSDTtoCustomXmlPart.doc");
        //ExEnd:BindSDTtoCustomXmlPart
    }

    void SetContentControlStyle()
    {
        //ExStart:SetContentControlStyle
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");

        auto sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        SharedPtr<Style> style = doc->get_Styles()->idx_get(StyleIdentifier::Quote);
        sdt->set_Style(style);

        doc->Save(ArtifactsDir + u"WorkingWithSdt.SetContentControlStyle.docx");
        //ExEnd:SetContentControlStyle
    }

    void CreatingTableRepeatingSectionMappedToCustomXmlPart()
    {
        //ExStart:CreatingTableRepeatingSectionMappedToCustomXmlPart
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<CustomXmlPart> xmlPart =
            doc->get_CustomXmlParts()->Add(u"Books", String(u"<books><book><title>Everyday Italian</title><author>Giada De Laurentiis</author></book>") +
                                                         u"<book><title>Harry Potter</title><author>J K. Rowling</author></book>" +
                                                         u"<book><title>Learning XML</title><author>Erik T. Ray</author></book></books>");

        SharedPtr<Table> table = builder->StartTable();

        builder->InsertCell();
        builder->Write(u"Title");

        builder->InsertCell();
        builder->Write(u"Author");

        builder->EndRow();
        builder->EndTable();

        auto repeatingSectionSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::RepeatingSection, MarkupLevel::Row);
        repeatingSectionSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book", u"");
        table->AppendChild(repeatingSectionSdt);

        auto repeatingSectionItemSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::RepeatingSectionItem, MarkupLevel::Row);
        repeatingSectionSdt->AppendChild(repeatingSectionItemSdt);

        auto row = MakeObject<Row>(doc);
        repeatingSectionItemSdt->AppendChild(row);

        auto titleSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Cell);
        titleSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/title[1]", u"");
        row->AppendChild(titleSdt);

        auto authorSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Cell);
        authorSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/author[1]", u"");
        row->AppendChild(authorSdt);

        doc->Save(ArtifactsDir + u"WorkingWithSdt.CreatingTableRepeatingSectionMappedToCustomXmlPart.docx");
        //ExEnd:CreatingTableRepeatingSectionMappedToCustomXmlPart
    }

    void MultiSection()
    {
        //ExStart:MultiSectionSDT
        auto doc = MakeObject<Document>(MyDir + u"Multi-section structured document tags.docx");

        SharedPtr<NodeCollection> tags = doc->GetChildNodes(NodeType::StructuredDocumentTagRangeStart, true);

        for (const auto& tag : System::IterateOver<StructuredDocumentTagRangeStart>(tags))
        {
            std::cout << tag->get_Title() << std::endl;
        }
        //ExEnd:MultiSectionSDT
    }

    void StructuredDocumentTagRangeStartXmlMapping()
    {
        //ExStart:StructuredDocumentTagRangeStartXmlMapping
        auto doc = MakeObject<Document>(MyDir + u"Multi-section structured document tags.docx");

        // Construct an XML part that contains data and add it to the document's CustomXmlPart collection.
        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Text element #1</text><text>Text element #2</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);
        std::cout << System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()) << std::endl;

        // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document.
        auto sdtRangeStart = System::DynamicCast<StructuredDocumentTagRangeStart>(doc->GetChild(NodeType::StructuredDocumentTagRangeStart, 0, true));

        // If we set a mapping for our StructuredDocumentTag,
        // it will only display a part of the CustomXmlPart that the XPath points to.
        // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
        sdtRangeStart->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[2]", nullptr);

        doc->Save(ArtifactsDir + u"WorkingWithSdt.StructuredDocumentTagRangeStartXmlMapping.docx");
        //ExEnd:StructuredDocumentTagRangeStartXmlMapping
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment
