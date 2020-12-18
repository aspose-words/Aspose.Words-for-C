#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlock.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPart.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPartCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlSchemaCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/MarkupLevel.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtCalendarType.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtDateStorageFormat.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtListItem.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtListItemCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtType.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagRangeStart.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/XmlMapping.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/guid.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

/// <summary>
/// Tests that verify work with structured document tags in the document.
/// </summary>
class ExStructuredDocumentTag : public ApiExampleBase
{
public:
    void RepeatingSection()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.SdtType
        //ExSummary:Shows how to get type of structured document tag.
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");

        SharedPtr<System::Collections::Generic::List<SharedPtr<StructuredDocumentTag>>> sdTags =
            doc->GetChildNodes(NodeType::StructuredDocumentTag, true)->LINQ_OfType<SharedPtr<StructuredDocumentTag>>()->LINQ_ToList();

        ASSERT_EQ(SdtType::RepeatingSection, sdTags->idx_get(0)->get_SdtType());
        ASSERT_EQ(SdtType::RepeatingSectionItem, sdTags->idx_get(1)->get_SdtType());
        ASSERT_EQ(SdtType::RichText, sdTags->idx_get(2)->get_SdtType());
        //ExEnd
    }

    void SetSpecificStyleToSdt()
    {
        //ExStart
        //ExFor:StructuredDocumentTag
        //ExFor:StructuredDocumentTag.NodeType
        //ExFor:StructuredDocumentTag.Style
        //ExFor:StructuredDocumentTag.StyleName
        //ExFor:MarkupLevel
        //ExFor:SdtType
        //ExSummary:Shows how to work with styles for content control elements.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Get specific style from the document to apply it to an SDT
        SharedPtr<Style> quoteStyle = doc->get_Styles()->idx_get(StyleIdentifier::Quote);
        auto sdtPlainText = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);
        sdtPlainText->set_Style(quoteStyle);

        auto sdtRichText = MakeObject<StructuredDocumentTag>(doc, SdtType::RichText, MarkupLevel::Inline);
        // Second method to apply specific style to an SDT control
        sdtRichText->set_StyleName(u"Quote");

        // Insert content controls into the document
        builder->InsertNode(sdtPlainText);
        builder->InsertNode(sdtRichText);

        // We can get a collection of StructuredDocumentTags by looking for the document's child nodes of this NodeType
        ASSERT_EQ(NodeType::StructuredDocumentTag, sdtPlainText->get_NodeType());

        SharedPtr<NodeCollection> tags = doc->GetChildNodes(NodeType::StructuredDocumentTag, true);

        for (auto node : System::IterateOver(tags))
        {
            auto sdt = System::DynamicCast<StructuredDocumentTag>(node);
            // If style was not defined before, style should be "Default Paragraph Font"
            ASSERT_EQ(StyleIdentifier::Quote, sdt->get_Style()->get_StyleIdentifier());
            ASSERT_EQ(u"Quote", sdt->get_StyleName());
        }
        //ExEnd
    }

    void CheckBox()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.#ctor(DocumentBase, SdtType, MarkupLevel)
        //ExFor:StructuredDocumentTag.Checked
        //ExSummary:Show how to create and insert checkbox structured document tag.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto sdtCheckBox = MakeObject<StructuredDocumentTag>(doc, SdtType::Checkbox, MarkupLevel::Inline);
        sdtCheckBox->set_Checked(true);

        // Insert content control into the document
        builder->InsertNode(sdtCheckBox);
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        SharedPtr<NodeCollection> sdts = doc->GetChildNodes(NodeType::StructuredDocumentTag, true);

        auto sdt = System::DynamicCast<StructuredDocumentTag>(sdts->idx_get(0));
        ASPOSE_ASSERT_EQ(true, sdt->get_Checked());
        ASSERT_EQ(sdt->get_XmlMapping()->get_StoreItemId(), String::Empty);
        // Assert that this sdt has no StoreItemId
    }

    void Date()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.CalendarType
        //ExFor:StructuredDocumentTag.DateDisplayFormat
        //ExFor:StructuredDocumentTag.DateDisplayLocale
        //ExFor:StructuredDocumentTag.DateStorageFormat
        //ExFor:StructuredDocumentTag.FullDate
        //ExSummary:Shows how to prompt the user to enter a date with a StructuredDocumentTag.
        // Create a new document
        auto doc = MakeObject<Document>();

        // Insert a StructuredDocumentTag that prompts the user to enter a date
        // In Microsoft Word, this element is known as a "Date picker content control"
        // When we click on the arrow on the right end of this tag in Microsoft Word,
        // we will see a pop up in the form of a clickable calendar
        // We can use that popup to select a date that will be displayed by the tag
        auto sdtDate = MakeObject<StructuredDocumentTag>(doc, SdtType::Date, MarkupLevel::Inline);

        // This attribute sets the language that the calendar will be displayed in,
        // which in this case will be Saudi Arabian Arabic
        sdtDate->set_DateDisplayLocale(System::Globalization::CultureInfo::GetCultureInfo(u"ar-SA")->get_LCID());

        // We can set the format with which to display the date like this
        // The locale we set above will be carried over to the displayed date
        sdtDate->set_DateDisplayFormat(u"dd MMMM, yyyy");

        // Select how the data will be stored in the document
        sdtDate->set_DateStorageFormat(SdtDateStorageFormat::DateTime);

        // Set the calendar type that will be used to select and display the date
        sdtDate->set_CalendarType(SdtCalendarType::Hijri);

        // Before a date is chosen, the tag will display the text "Click here to enter a date."
        // We can set a default date to display by setting this variable
        // We must convert the date to the appropriate calendar ourselves
        sdtDate->set_FullDate(System::DateTime(1440, 10, 20));

        // Insert the StructuredDocumentTag into the document with a DocumentBuilder and save the document
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertNode(sdtDate);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.Date.docx");
        //ExEnd
    }

    void PlainText()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.Color
        //ExFor:StructuredDocumentTag.ContentsFont
        //ExFor:StructuredDocumentTag.EndCharacterFont
        //ExFor:StructuredDocumentTag.Id
        //ExFor:StructuredDocumentTag.Level
        //ExFor:StructuredDocumentTag.Multiline
        //ExFor:StructuredDocumentTag.Tag
        //ExFor:StructuredDocumentTag.Title
        //ExFor:StructuredDocumentTag.RemoveSelfOnly
        //ExSummary:Shows how to create a StructuredDocumentTag in the form of a plain text box and modify its appearance.
        // Create a new document
        auto doc = MakeObject<Document>();

        // Create a StructuredDocumentTag that will contain plain text
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // Set the title and color of the frame that appears when you mouse over it
        tag->set_Title(u"My plain text");
        tag->set_Color(System::Drawing::Color::get_Magenta());

        // Set a programmatic tag for this StructuredDocumentTag
        // Unlike the title, this value will not be visible in the document but will be programmatically obtainable
        // as an XML element named "tag", with the string below in its "@val" attribute
        tag->set_Tag(u"MyPlainTextSDT");

        // Every StructuredDocumentTag gets a random unique ID
        ASSERT_GT(tag->get_Id(), 0);

        // Set the font for the text inside the StructuredDocumentTag
        tag->get_ContentsFont()->set_Name(u"Arial");

        // Set the font for the text at the end of the StructuredDocumentTag
        // Any text that is typed in the document body after moving out of the tag with arrow keys will keep this font
        tag->get_EndCharacterFont()->set_Name(u"Arial Black");

        // By default, this is false and pressing enter while inside a StructuredDocumentTag does nothing
        // When set to true, our StructuredDocumentTag can have multiple lines
        tag->set_Multiline(true);

        // Insert the StructuredDocumentTag into the document with a DocumentBuilder and save the document to a file
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertNode(tag);

        // Insert a clone of our StructuredDocumentTag in a new paragraph
        auto tagClone = System::DynamicCast<StructuredDocumentTag>(tag->Clone(true));
        builder->InsertParagraph();
        builder->InsertNode(tagClone);

        // We can remove the tag while keeping its contents where they were in the Paragraph by calling RemoveSelfOnly()
        tagClone->RemoveSelfOnly();

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.PlainText.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.PlainText.docx");
        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));

        ASSERT_EQ(u"My plain text", tag->get_Title());
        ASSERT_EQ(System::Drawing::Color::get_Magenta().ToArgb(), tag->get_Color().ToArgb());
        ASSERT_EQ(u"MyPlainTextSDT", tag->get_Tag());
        ASSERT_TRUE(tag->get_Id() > 0);
        ASSERT_EQ(u"Arial", tag->get_ContentsFont()->get_Name());
        ASSERT_EQ(u"Arial Black", tag->get_EndCharacterFont()->get_Name());
        ASSERT_TRUE(tag->get_Multiline());
    }

    void IsTemporary()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsTemporary
        //ExSummary:Demonstrates the effects of making a StructuredDocumentTag temporary.
        auto doc = MakeObject<Document>();

        // Insert a plain text StructuredDocumentTag, which will prompt the user to enter text
        // and allow them to edit it like a text box
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // If we set its Temporary attribute to true, as soon as we start typing,
        // the tag will disappear, and its contents will be assimilated into the parent Paragraph
        tag->set_IsTemporary(true);

        // Insert the StructuredDocumentTag with a DocumentBuilder
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Temporary text box: ");
        builder->InsertNode(tag);

        // A StructuredDocumentTag in the form of a check box will let the user a square to check and uncheck
        // Setting it to temporary will freeze its value after the first time it is clicked
        tag = MakeObject<StructuredDocumentTag>(doc, SdtType::Checkbox, MarkupLevel::Inline);
        tag->set_IsTemporary(true);

        builder->Write(u"\nTemporary checkbox: ");
        builder->InsertNode(tag);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.IsTemporary.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.IsTemporary.docx");

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> sdt)> _local_func_0 = [](SharedPtr<Node> sdt)
        {
            return (System::DynamicCast<StructuredDocumentTag>(sdt))->get_IsTemporary();
        };

        ASSERT_EQ(
            2, doc->GetChildNodes(NodeType::StructuredDocumentTag, true)->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_0)));
    }

    void PlaceholderBuildingBlock()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsShowingPlaceholderText
        //ExFor:StructuredDocumentTag.Placeholder
        //ExFor:StructuredDocumentTag.PlaceholderName
        //ExSummary:Shows how to use the contents of a BuildingBlock as a custom placeholder text for a StructuredDocumentTag.
        auto doc = MakeObject<Document>();

        // Insert a plain text StructuredDocumentTag of the PlainText type, which will function like a text box
        // It contains a default "Click here to enter text." prompt, which we can click and replace with our own text
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // We can substitute that default placeholder with a custom phrase, which will be drawn from a BuildingBlock
        // First, we will need to create the BuildingBlock, give it content and add it to the GlossaryDocument
        SharedPtr<GlossaryDocument> glossaryDoc = doc->get_GlossaryDocument();

        auto substituteBlock = MakeObject<BuildingBlock>(glossaryDoc);
        substituteBlock->set_Name(u"Custom Placeholder");
        substituteBlock->AppendChild(MakeObject<Section>(glossaryDoc));
        substituteBlock->get_FirstSection()->AppendChild(MakeObject<Body>(glossaryDoc));
        substituteBlock->get_FirstSection()->get_Body()->AppendParagraph(u"Custom placeholder text.");

        glossaryDoc->AppendChild(substituteBlock);

        // The substitute BuildingBlock we made can be referenced by name
        tag->set_PlaceholderName(u"Custom Placeholder");

        // If PlaceholderName refers to an existing block in the parent document's GlossaryDocument,
        // the BuildingBlock will be automatically found and assigned to the Placeholder attribute
        ASPOSE_ASSERT_EQ(substituteBlock, tag->get_Placeholder());

        // Setting this to true will register the text inside the StructuredDocumentTag as placeholder text
        // This means that, in Microsoft Word, all the text contents of the StructuredDocumentTag will be highlighted with one click,
        // so we can immediately replace the entire substitute text by typing
        // If this is false, the text will behave like an ordinary Paragraph and a cursor will be placed with nothing highlighted
        tag->set_IsShowingPlaceholderText(true);

        // Insert the StructuredDocumentTag into the document using a DocumentBuilder and save the document to a file
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertNode(tag);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        substituteBlock = System::DynamicCast<BuildingBlock>(doc->get_GlossaryDocument()->GetChild(NodeType::BuildingBlock, 0, true));

        ASSERT_EQ(u"Custom Placeholder", substituteBlock->get_Name());
        ASSERT_TRUE(tag->get_IsShowingPlaceholderText());
        ASPOSE_ASSERT_EQ(substituteBlock, tag->get_Placeholder());
        ASSERT_EQ(substituteBlock->get_Name(), tag->get_PlaceholderName());
    }

    void Lock()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.LockContentControl
        //ExFor:StructuredDocumentTag.LockContents
        //ExSummary:Shows how to restrict the editing of a StructuredDocumentTag.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a plain text StructuredDocumentTag of the PlainText type, which will function like a text box
        // It contains a default "Click here to enter text." prompt, which we can click and replace with our own text
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // We can prohibit the users from editing the inner text in Microsoft Word by setting this to true
        tag->set_LockContents(true);
        builder->Write(u"The contents of this StructuredDocumentTag cannot be edited: ");
        builder->InsertNode(tag);

        tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // Setting this to true will disable the deletion of this StructuredDocumentTag
        // by text editing operations in Microsoft Word
        tag->set_LockContentControl(true);

        builder->InsertParagraph();
        builder->Write(u"This StructuredDocumentTag cannot be deleted but its contents can be edited: ");
        builder->InsertNode(tag);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.Lock.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.Lock.docx");
        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));

        ASSERT_TRUE(tag->get_LockContents());
        ASSERT_FALSE(tag->get_LockContentControl());

        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 1, true));

        ASSERT_FALSE(tag->get_LockContents());
        ASSERT_TRUE(tag->get_LockContentControl());
    }

    void ListItemCollection()
    {
        //ExStart
        //ExFor:SdtListItem
        //ExFor:SdtListItem.#ctor(System.String)
        //ExFor:SdtListItem.#ctor(System.String,System.String)
        //ExFor:SdtListItem.DisplayText
        //ExFor:SdtListItem.Value
        //ExFor:SdtListItemCollection
        //ExFor:SdtListItemCollection.Add(Aspose.Words.Markup.SdtListItem)
        //ExFor:SdtListItemCollection.Clear
        //ExFor:SdtListItemCollection.Count
        //ExFor:SdtListItemCollection.GetEnumerator
        //ExFor:SdtListItemCollection.Item(System.Int32)
        //ExFor:SdtListItemCollection.RemoveAt(System.Int32)
        //ExFor:SdtListItemCollection.SelectedValue
        //ExFor:StructuredDocumentTag.ListItems
        //ExSummary:Shows how to work with StructuredDocumentTag nodes of the DropDownList type.
        // Create a blank document and insert a StructuredDocumentTag that will contain a drop-down list
        auto doc = MakeObject<Document>();
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::DropDownList, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        // A drop-down list needs elements, each of which will be a SdtListItem
        SharedPtr<SdtListItemCollection> listItems = tag->get_ListItems();
        listItems->Add(MakeObject<SdtListItem>(u"Value 1"));

        // Each SdtListItem has text that will be displayed when the drop-down list is opened, and also a value
        // When we initialize with one string, we are providing just the value
        // Accordingly, value is passed as DisplayText and will consequently be displayed on the screen
        ASSERT_EQ(listItems->idx_get(0)->get_DisplayText(), listItems->idx_get(0)->get_Value());

        // Add 3 more SdtListItems with non-empty strings passed to DisplayText
        listItems->Add(MakeObject<SdtListItem>(u"Item 2", u"Value 2"));
        listItems->Add(MakeObject<SdtListItem>(u"Item 3", u"Value 3"));
        listItems->Add(MakeObject<SdtListItem>(u"Item 4", u"Value 4"));

        // We can obtain a count of the SdtListItems and also set the drop-down list's SelectedValue attribute to
        // automatically have one of them pre-selected when we open the document in Microsoft Word
        ASSERT_EQ(4, listItems->get_Count());
        listItems->set_SelectedValue(listItems->idx_get(3));

        ASSERT_EQ(u"Value 4", listItems->get_SelectedValue()->get_Value());

        // We can enumerate over the collection and print each element
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<SdtListItem>>> enumerator = listItems->GetEnumerator();
            while (enumerator->MoveNext())
            {
                if (enumerator->get_Current() != nullptr)
                {
                    std::cout << "List item: " << enumerator->get_Current()->get_DisplayText() << ", value: " << enumerator->get_Current()->get_Value()
                              << std::endl;
                }
            }
        }

        // We can also remove elements one at a time
        listItems->RemoveAt(3);
        ASSERT_EQ(3, listItems->get_Count());

        // Make sure to update the SelectedValue's index if it ever ends up out of bounds before saving the document
        listItems->set_SelectedValue(listItems->idx_get(1));

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.ListItemCollection.docx");

        // We can clear the whole collection at once too
        listItems->Clear();
        ASSERT_EQ(0, listItems->get_Count());
        //ExEnd
    }

    void CreatingCustomXml()
    {
        //ExStart
        //ExFor:CustomXmlPart
        //ExFor:CustomXmlPart.Clone
        //ExFor:CustomXmlPart.Data
        //ExFor:CustomXmlPart.Id
        //ExFor:CustomXmlPart.Schemas
        //ExFor:CustomXmlPartCollection
        //ExFor:CustomXmlPartCollection.Add(CustomXmlPart)
        //ExFor:CustomXmlPartCollection.Add(String, String)
        //ExFor:CustomXmlPartCollection.Clear
        //ExFor:CustomXmlPartCollection.Clone
        //ExFor:CustomXmlPartCollection.Count
        //ExFor:CustomXmlPartCollection.GetById(String)
        //ExFor:CustomXmlPartCollection.GetEnumerator
        //ExFor:CustomXmlPartCollection.Item(Int32)
        //ExFor:CustomXmlPartCollection.RemoveAt(Int32)
        //ExFor:Document.CustomXmlParts
        //ExFor:StructuredDocumentTag.XmlMapping
        //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
        //ExSummary:Shows how to create structured document tag with a custom XML data.
        auto doc = MakeObject<Document>();

        // Construct an XML part that contains data and add it to the document's collection
        // Once the "Developer" tab in Microsoft Word is enabled,
        // we can find elements from this collection as well as a couple defaults in the "XML Mapping Pane"
        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Hello world!</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);

        // The data we entered is stored in these attributes
        ASPOSE_ASSERT_EQ(System::Text::Encoding::get_ASCII()->GetBytes(xmlPartContent), xmlPart->get_Data());
        ASSERT_EQ(xmlPartId, xmlPart->get_Id());

        // XML parts can be referenced by collection index or GUID
        ASPOSE_ASSERT_EQ(xmlPart, doc->get_CustomXmlParts()->idx_get(0));
        ASPOSE_ASSERT_EQ(xmlPart, doc->get_CustomXmlParts()->GetById(xmlPartId));

        // Once the part is created, we can add XML schema associations like this
        xmlPart->get_Schemas()->Add(u"http://www.w3.org/2001/XMLSchema");

        // We can also clone parts and insert them into the collection directly
        SharedPtr<CustomXmlPart> xmlPartClone = xmlPart->Clone();
        xmlPartClone->set_Id(System::Guid::NewGuid().ToString(u"B"));
        doc->get_CustomXmlParts()->Add(xmlPartClone);

        ASSERT_EQ(2, doc->get_CustomXmlParts()->get_Count());

        // Iterate through collection with an enumerator and print the contents of each part
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<CustomXmlPart>>> enumerator =
                doc->get_CustomXmlParts()->GetEnumerator();
            int index = 0;
            while (enumerator->MoveNext())
            {
                std::cout << "XML part index " << index << ", ID: " << enumerator->get_Current()->get_Id() << std::endl;
                std::cout << "\tContent: " << System::Text::Encoding::get_UTF8()->GetString(enumerator->get_Current()->get_Data()) << std::endl;
                index++;
            }
        }

        // XML parts can be removed by index
        doc->get_CustomXmlParts()->RemoveAt(1);

        ASSERT_EQ(1, doc->get_CustomXmlParts()->get_Count());

        // The XML part collection itself can be cloned also
        SharedPtr<CustomXmlPartCollection> customXmlParts = doc->get_CustomXmlParts()->Clone();

        // And all elements can be cleared like this
        customXmlParts->Clear();

        // Create a StructuredDocumentTag that will display the contents of our part,
        // insert it into the document and save the document
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        tag->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[1]", String::Empty);

        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.CustomXml.docx");
        //ExEnd

        ASSERT_TRUE(
            DocumentHelper::CompareDocs(ArtifactsDir + u"StructuredDocumentTag.CustomXml.docx", GoldsDir + u"StructuredDocumentTag.CustomXml Gold.docx"));

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.CustomXml.docx");
        xmlPart = doc->get_CustomXmlParts()->idx_get(0);

        System::Guid temp;
        ASSERT_TRUE(System::Guid::TryParse(xmlPart->get_Id(), temp));
        ASSERT_EQ(u"<root><text>Hello world!</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));
        ASSERT_EQ(u"http://www.w3.org/2001/XMLSchema", xmlPart->get_Schemas()->idx_get(0));

        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        ASSERT_EQ(u"Hello world!", tag->GetText().Trim());
        ASSERT_EQ(u"/root[1]/text[1]", tag->get_XmlMapping()->get_XPath());
        ASSERT_EQ(String::Empty, tag->get_XmlMapping()->get_PrefixMappings());
    }

    void XmlMapping_()
    {
        //ExStart
        //ExFor:XmlMapping
        //ExFor:XmlMapping.CustomXmlPart
        //ExFor:XmlMapping.Delete
        //ExFor:XmlMapping.IsMapped
        //ExFor:XmlMapping.PrefixMappings
        //ExFor:XmlMapping.XPath
        //ExSummary:Shows how to set XML mappings for CustomXmlParts.
        auto doc = MakeObject<Document>();

        // Construct an XML part that contains data and add it to the document's CustomXmlPart collection
        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Text element #1</text><text>Text element #2</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);
        std::cout << System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()) << std::endl;

        // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);

        // If we set a mapping for our StructuredDocumentTag,
        // it will only display a part of the CustomXmlPart that the XPath points to
        // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart
        tag->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[2]", u"xmlns:ns='http://www.w3.org/2001/XMLSchema'");

        ASSERT_TRUE(tag->get_XmlMapping()->get_IsMapped());
        ASPOSE_ASSERT_EQ(xmlPart, tag->get_XmlMapping()->get_CustomXmlPart());
        ASSERT_EQ(u"/root[1]/text[2]", tag->get_XmlMapping()->get_XPath());
        ASSERT_EQ(u"xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag->get_XmlMapping()->get_PrefixMappings());

        // Add the StructuredDocumentTag to the document to display the content from our CustomXmlPart
        doc->get_FirstSection()->get_Body()->AppendChild(tag);
        doc->Save(ArtifactsDir + u"StructuredDocumentTag.XmlMapping.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.XmlMapping.docx");
        xmlPart = doc->get_CustomXmlParts()->idx_get(0);

        System::Guid temp;
        ASSERT_TRUE(System::Guid::TryParse(xmlPart->get_Id(), temp));
        ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));

        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        ASSERT_EQ(u"Text element #2", tag->GetText().Trim());
        ASSERT_EQ(u"/root[1]/text[2]", tag->get_XmlMapping()->get_XPath());
        ASSERT_EQ(u"xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag->get_XmlMapping()->get_PrefixMappings());
    }

    void CustomXmlSchemaCollection_()
    {
        //ExStart
        //ExFor:CustomXmlSchemaCollection
        //ExFor:CustomXmlSchemaCollection.Add(System.String)
        //ExFor:CustomXmlSchemaCollection.Clear
        //ExFor:CustomXmlSchemaCollection.Clone
        //ExFor:CustomXmlSchemaCollection.Count
        //ExFor:CustomXmlSchemaCollection.GetEnumerator
        //ExFor:CustomXmlSchemaCollection.IndexOf(System.String)
        //ExFor:CustomXmlSchemaCollection.Item(System.Int32)
        //ExFor:CustomXmlSchemaCollection.Remove(System.String)
        //ExFor:CustomXmlSchemaCollection.RemoveAt(System.Int32)
        //ExSummary:Shows how to work with an XML schema collection.
        // Create a document and add a custom XML part
        auto doc = MakeObject<Document>();

        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Hello, World!</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);

        // Once the part is created, we can add XML schema associations like this,
        // and perform other collection-related operations on the list of schemas for this part
        xmlPart->get_Schemas()->Add(u"http://www.w3.org/2001/XMLSchema");

        // Collections can be cloned, and elements can be added
        SharedPtr<CustomXmlSchemaCollection> schemas = xmlPart->get_Schemas()->Clone();
        schemas->Add(u"http://www.w3.org/2001/XMLSchema-instance");
        schemas->Add(u"http://schemas.microsoft.com/office/2006/metadata/contentType");

        ASSERT_EQ(3, schemas->get_Count());
        ASSERT_EQ(2, schemas->IndexOf((u"http://schemas.microsoft.com/office/2006/metadata/contentType")));

        // We can iterate over the collection with an enumerator
        {
            SharedPtr<System::Collections::Generic::IEnumerator<String>> enumerator = schemas->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << enumerator->get_Current() << std::endl;
            }
        }

        // We can also remove elements by index, element, or we can clear the entire collection
        schemas->RemoveAt(2);
        schemas->Remove(u"http://www.w3.org/2001/XMLSchema");
        schemas->Clear();

        ASSERT_EQ(0, schemas->get_Count());
        //ExEnd
    }

    void CustomXmlPartStoreItemIdReadOnly()
    {
        //ExStart
        //ExFor:XmlMapping.StoreItemId
        //ExSummary:Shows how to get special id of your xml part.
        auto doc = MakeObject<Document>(MyDir + u"Custom XML part in structured document tag.docx");

        // Structured document tags have IDs in the form of Guids
        auto tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        ASSERT_EQ(u"{F3029283-4FF8-4DD2-9F31-395F19ACEE85}", tag->get_XmlMapping()->get_StoreItemId());
        //ExEnd
    }

    void CustomXmlPartStoreItemIdReadOnlyNull()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto sdtCheckBox = MakeObject<StructuredDocumentTag>(doc, SdtType::Checkbox, MarkupLevel::Inline);
        sdtCheckBox->set_Checked(true);

        // Insert content control into the document
        builder->InsertNode(sdtCheckBox);

        doc = DocumentHelper::SaveOpen(doc);

        auto sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        std::cout << (String(u"The Id of your custom xml part is: ") + sdt->get_XmlMapping()->get_StoreItemId()) << std::endl;
    }

    void ClearTextFromStructuredDocumentTags()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.Clear
        //ExSummary:Shows how to delete content of StructuredDocumentTag elements.
        auto doc = MakeObject<Document>();

        // Create a plain text structured document tag and append it to the document
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        // This structured document tag, which is in the form of a text box, already displays placeholder text
        ASSERT_EQ(u"Click here to enter text.", tag->GetText().Trim());
        ASSERT_TRUE(tag->get_IsShowingPlaceholderText());

        // Create a building block that
        SharedPtr<GlossaryDocument> glossaryDoc = doc->get_GlossaryDocument();
        auto substituteBlock = MakeObject<BuildingBlock>(glossaryDoc);
        substituteBlock->set_Name(u"My placeholder");
        substituteBlock->AppendChild(MakeObject<Section>(glossaryDoc));
        substituteBlock->get_FirstSection()->EnsureMinimum();
        substituteBlock->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(glossaryDoc, u"Custom placeholder text."));
        glossaryDoc->AppendChild(substituteBlock);

        // Set the tag's placeholder to the building block
        tag->set_PlaceholderName(u"My placeholder");

        ASSERT_EQ(u"Custom placeholder text.", tag->GetText().Trim());
        ASSERT_TRUE(tag->get_IsShowingPlaceholderText());

        // Edit the text of the structured document tag and disable showing of placeholder text
        auto run = System::DynamicCast<Run>(tag->GetChild(NodeType::Run, 0, true));
        run->set_Text(u"New text.");
        tag->set_IsShowingPlaceholderText(false);

        ASSERT_EQ(u"New text.", tag->GetText().Trim());

        tag->Clear();

        // Clearing a PlainText tag reverts these changes
        ASSERT_TRUE(tag->get_IsShowingPlaceholderText());
        ASSERT_EQ(u"Custom placeholder text.", tag->GetText().Trim());
        //ExEnd
    }

    void AccessToBuildingBlockPropertiesFromDocPartObjSdt()
    {
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags with building blocks.docx");

        auto docPartObjSdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));

        ASSERT_EQ(SdtType::DocPartObj, docPartObjSdt->get_SdtType());
        ASSERT_EQ(u"Table of Contents", docPartObjSdt->get_BuildingBlockGallery());
    }

    void AccessToBuildingBlockPropertiesFromPlainTextSdt()
    {
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags with building blocks.docx");

        auto plainTextSdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 1, true));

        ASSERT_EQ(SdtType::PlainText, plainTextSdt->get_SdtType());

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<String()> _local_func_1 = [&plainTextSdt]()
        {
            return plainTextSdt->get_BuildingBlockGallery();
        };

        ASSERT_THROW(_local_func_1(), System::InvalidOperationException);
    }

    void BuildingBlockCategories()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.BuildingBlockCategory
        //ExFor:StructuredDocumentTag.BuildingBlockGallery
        //ExSummary:Shows how to insert a StructuredDocumentTag as a building block and set its category and gallery.
        auto doc = MakeObject<Document>();

        auto buildingBlockSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::BuildingBlockGallery, MarkupLevel::Block);
        buildingBlockSdt->set_BuildingBlockCategory(u"Built-in");
        buildingBlockSdt->set_BuildingBlockGallery(u"Table of Contents");

        doc->get_FirstSection()->get_Body()->AppendChild(buildingBlockSdt);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.BuildingBlockCategories.docx");
        //ExEnd

        buildingBlockSdt = System::DynamicCast<StructuredDocumentTag>(doc->get_FirstSection()->get_Body()->GetChild(NodeType::StructuredDocumentTag, 0, true));

        ASSERT_EQ(SdtType::BuildingBlockGallery, buildingBlockSdt->get_SdtType());
        ASSERT_EQ(u"Table of Contents", buildingBlockSdt->get_BuildingBlockGallery());
        ASSERT_EQ(u"Built-in", buildingBlockSdt->get_BuildingBlockCategory());
    }

    void UpdateSdtContent(bool updateSdtContent)
    {
        //ExStart
        //ExFor:SaveOptions.UpdateSdtContent
        //ExSummary:Shows how structured document tags can be updated while saving to .pdf.
        auto doc = MakeObject<Document>();

        // Insert two StructuredDocumentTags; a date and a drop-down list
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::Date, MarkupLevel::Block);
        tag->set_FullDate(System::DateTime::get_Now());

        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        tag = MakeObject<StructuredDocumentTag>(doc, SdtType::DropDownList, MarkupLevel::Block);
        tag->get_ListItems()->Add(MakeObject<SdtListItem>(u"Value 1"));
        tag->get_ListItems()->Add(MakeObject<SdtListItem>(u"Value 2"));
        tag->get_ListItems()->Add(MakeObject<SdtListItem>(u"Value 3"));
        tag->get_ListItems()->set_SelectedValue(tag->get_ListItems()->idx_get(1));

        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        // We've selected default values for both tags
        // We can save those values in the document without immediately updating the tags, leaving them in their default state
        // by using a SaveOptions object with this flag set
        auto options = MakeObject<PdfSaveOptions>();
        options->set_UpdateSdtContent(updateSdtContent);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.UpdateSdtContent.pdf", options);
        //ExEnd

    }

    void FillTableUsingRepeatingSectionItem()
    {
        //ExStart
        //ExFor:SdtType
        //ExSummary:Shows how to fill the table with data contained in the XML part.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(
            u"Books", String(u"<books>") + u"<book><title>Everyday Italian</title>" + u"<author>Giada De Laurentiis</author></book>" +
                          u"<book><title>Harry Potter</title>" + u"<author>J. K. Rowling</author></book>" + u"<book><title>Learning XML</title>" +
                          u"<author>Erik T. Ray</author></book>" + u"</books>");

        // Create headers for data from xml content
        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Title");
        builder->InsertCell();
        builder->Write(u"Author");
        builder->EndRow();
        builder->EndTable();

        // Create table with RepeatingSection inside
        auto repeatingSectionSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::RepeatingSection, MarkupLevel::Row);
        repeatingSectionSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book", String::Empty);
        table->AppendChild(repeatingSectionSdt);

        // Add RepeatingSectionItem inside RepeatingSection and mark it as a row
        auto repeatingSectionItemSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::RepeatingSectionItem, MarkupLevel::Row);
        repeatingSectionSdt->AppendChild(repeatingSectionItemSdt);

        auto row = MakeObject<Row>(doc);
        repeatingSectionItemSdt->AppendChild(row);

        // Map xml data with created table cells for book title and author
        auto titleSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Cell);
        titleSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/title[1]", String::Empty);
        row->AppendChild(titleSdt);

        auto authorSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Cell);
        authorSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/author[1]", String::Empty);
        row->AppendChild(authorSdt);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.RepeatingSectionItem.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.RepeatingSectionItem.docx");
        SharedPtr<System::Collections::Generic::List<SharedPtr<StructuredDocumentTag>>> tags =
            doc->GetChildNodes(NodeType::StructuredDocumentTag, true)->LINQ_OfType<SharedPtr<StructuredDocumentTag>>()->LINQ_ToList();

        ASSERT_EQ(u"/books[1]/book", tags->idx_get(0)->get_XmlMapping()->get_XPath());
        ASSERT_EQ(String::Empty, tags->idx_get(0)->get_XmlMapping()->get_PrefixMappings());

        ASSERT_EQ(String::Empty, tags->idx_get(1)->get_XmlMapping()->get_XPath());
        ASSERT_EQ(String::Empty, tags->idx_get(1)->get_XmlMapping()->get_PrefixMappings());

        ASSERT_EQ(u"/books[1]/book[1]/title[1]", tags->idx_get(2)->get_XmlMapping()->get_XPath());
        ASSERT_EQ(String::Empty, tags->idx_get(2)->get_XmlMapping()->get_PrefixMappings());

        ASSERT_EQ(u"/books[1]/book[1]/author[1]", tags->idx_get(3)->get_XmlMapping()->get_XPath());
        ASSERT_EQ(String::Empty, tags->idx_get(3)->get_XmlMapping()->get_PrefixMappings());

        ASSERT_EQ(String(u"Title\u0007Author\u0007\u0007") + u"Everyday Italian\u0007Giada De Laurentiis\u0007\u0007" +
                      u"Harry Potter\u0007J. K. Rowling\u0007\u0007" + u"Learning XML\u0007Erik T. Ray\u0007\u0007",
                  doc->GetChild(NodeType::Table, 0, true)->GetText().Trim());
    }

    void CustomXmlPart_()
    {
        // Obtain an XML in the form of a string
        String xmlString = String(u"<?xml version=\"1.0\"?>") + u"<Company>" + u"<Employee id=\"1\">" + u"<FirstName>John</FirstName>" +
                                   u"<LastName>Doe</LastName>" + u"</Employee>" + u"<Employee id=\"2\">" + u"<FirstName>Jane</FirstName>" +
                                   u"<LastName>Doe</LastName>" + u"</Employee>" + u"</Company>";

        auto doc = MakeObject<Document>();

        // Insert the full XML document as a custom document part
        // The mapping for this part will be seen in the "XML Mapping Pane" in the "Developer" tab, if it is enabled
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(System::Guid::NewGuid().ToString(u"B"), xmlString);

        // None of the XML is in the document body at this point
        // Create a StructuredDocumentTag, which will refer to a single element from the XML with an XPath
        auto sdt = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        sdt->get_XmlMapping()->SetMapping(xmlPart, u"Company//Employee[@id='2']/FirstName", u"");

        // Add the StructuredDocumentTag to the document to display the element in the text
        doc->get_FirstSection()->get_Body()->AppendChild(sdt);
    }

    void MultiSectionTags()
    {
        //ExStart
        //ExFor:StructuredDocumentTagRangeStart
        //ExFor:StructuredDocumentTagRangeStart.Id
        //ExFor:StructuredDocumentTagRangeStart.Title
        //ExFor:StructuredDocumentTagRangeStart.IsShowingPlaceholderText
        //ExFor:StructuredDocumentTagRangeStart.LockContentControl
        //ExFor:StructuredDocumentTagRangeStart.LockContents
        //ExFor:StructuredDocumentTagRangeStart.Level
        //ExFor:StructuredDocumentTagRangeStart.RangeEnd
        //ExFor:StructuredDocumentTagRangeStart.SdtType
        //ExFor:StructuredDocumentTagRangeStart.Tag
        //ExFor:StructuredDocumentTagRangeEnd
        //ExFor:StructuredDocumentTagRangeEnd.Id
        //ExSummary:Shows how to get multi-section structured document tags properties.
        auto doc = MakeObject<Document>(MyDir + u"Multi-section structured document tags.docx");

        // Note that these nodes can be a child of NodeType.Body node only and all properties of these nodes are read-only.
        auto rangeStartTag =
            System::DynamicCast_noexcept<StructuredDocumentTagRangeStart>(doc->GetChildNodes(NodeType::StructuredDocumentTagRangeStart, true)->idx_get(0));
        auto rangeEndTag =
            System::DynamicCast_noexcept<StructuredDocumentTagRangeEnd>(doc->GetChildNodes(NodeType::StructuredDocumentTagRangeEnd, true)->idx_get(0));

        ASSERT_EQ(rangeStartTag->get_Id(), rangeEndTag->get_Id());
        //ExSkip
        ASSERT_EQ(NodeType::StructuredDocumentTagRangeStart, rangeStartTag->get_NodeType());
        //ExSkip
        ASSERT_EQ(NodeType::StructuredDocumentTagRangeEnd, rangeEndTag->get_NodeType());
        //ExSkip

        std::cout << "StructuredDocumentTagRangeStart values:" << std::endl;
        std::cout << "\t|Id: " << rangeStartTag->get_Id() << std::endl;
        std::cout << "\t|Title: " << rangeStartTag->get_Title() << std::endl;
        std::cout << String::Format(u"\t|IsShowingPlaceholderText: {0}", rangeStartTag->get_IsShowingPlaceholderText()) << std::endl;
        std::cout << String::Format(u"\t|LockContentControl: {0}", rangeStartTag->get_LockContentControl()) << std::endl;
        std::cout << String::Format(u"\t|LockContents: {0}", rangeStartTag->get_LockContents()) << std::endl;
        std::cout << String::Format(u"\t|Level: {0}", rangeStartTag->get_Level()) << std::endl;
        std::cout << String::Format(u"\t|NodeType: {0}", rangeStartTag->get_NodeType()) << std::endl;
        std::cout << "\t|RangeEnd: " << rangeStartTag->get_RangeEnd() << std::endl;
        std::cout << String::Format(u"\t|SdtType: {0}", rangeStartTag->get_SdtType()) << std::endl;
        std::cout << "\t|Tag: " << rangeStartTag->get_Tag() << "\n" << std::endl;

        std::cout << "StructuredDocumentTagRangeEnd values:" << std::endl;
        std::cout << "\t|Id: " << rangeEndTag->get_Id() << std::endl;
        std::cout << String::Format(u"\t|NodeType: {0}", rangeEndTag->get_NodeType()) << std::endl;
        //ExEnd
    }
};

} // namespace ApiExamples
