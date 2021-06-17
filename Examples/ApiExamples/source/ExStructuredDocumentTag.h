#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BuildingBlocks/BuildingBlock.h>
#include <Aspose.Words.Cpp/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Markup/CustomXmlPart.h>
#include <Aspose.Words.Cpp/Markup/CustomXmlPartCollection.h>
#include <Aspose.Words.Cpp/Markup/CustomXmlSchemaCollection.h>
#include <Aspose.Words.Cpp/Markup/MarkupLevel.h>
#include <Aspose.Words.Cpp/Markup/SdtCalendarType.h>
#include <Aspose.Words.Cpp/Markup/SdtDateStorageFormat.h>
#include <Aspose.Words.Cpp/Markup/SdtListItem.h>
#include <Aspose.Words.Cpp/Markup/SdtListItemCollection.h>
#include <Aspose.Words.Cpp/Markup/SdtType.h>
#include <Aspose.Words.Cpp/Markup/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Markup/StructuredDocumentTagRangeEnd.h>
#include <Aspose.Words.Cpp/Markup/StructuredDocumentTagRangeStart.h>
#include <Aspose.Words.Cpp/Markup/XmlMapping.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <drawing/color.h>
#include <system/array.h>
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
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExStructuredDocumentTag : public ApiExampleBase
{
public:
    void RepeatingSection()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.SdtType
        //ExSummary:Shows how to get the type of a structured document tag.
        auto doc = MakeObject<Document>(MyDir + u"Structured document tags.docx");

        SharedPtr<System::Collections::Generic::List<SharedPtr<StructuredDocumentTag>>> sdTags =
            doc->GetChildNodes(NodeType::StructuredDocumentTag, true)->LINQ_OfType<SharedPtr<StructuredDocumentTag>>()->LINQ_ToList();

        ASSERT_EQ(SdtType::RepeatingSection, sdTags->idx_get(0)->get_SdtType());
        ASSERT_EQ(SdtType::RepeatingSectionItem, sdTags->idx_get(1)->get_SdtType());
        ASSERT_EQ(SdtType::RichText, sdTags->idx_get(2)->get_SdtType());
        //ExEnd
    }

    void ApplyStyle()
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

        // Below are two ways to apply a style from the document to a structured document tag.
        // 1 -  Apply a style object from the document's style collection:
        SharedPtr<Style> quoteStyle = doc->get_Styles()->idx_get(StyleIdentifier::Quote);
        auto sdtPlainText = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);
        sdtPlainText->set_Style(quoteStyle);

        // 2 -  Reference a style in the document by name:
        auto sdtRichText = MakeObject<StructuredDocumentTag>(doc, SdtType::RichText, MarkupLevel::Inline);
        sdtRichText->set_StyleName(u"Quote");

        builder->InsertNode(sdtPlainText);
        builder->InsertNode(sdtRichText);

        ASSERT_EQ(NodeType::StructuredDocumentTag, sdtPlainText->get_NodeType());

        SharedPtr<NodeCollection> tags = doc->GetChildNodes(NodeType::StructuredDocumentTag, true);

        for (const auto& node : System::IterateOver(tags))
        {
            auto sdt = System::DynamicCast<StructuredDocumentTag>(node);

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
        //ExSummary:Show how to create a structured document tag in the form of a check box.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto sdtCheckBox = MakeObject<StructuredDocumentTag>(doc, SdtType::Checkbox, MarkupLevel::Inline);
        sdtCheckBox->set_Checked(true);

        builder->InsertNode(sdtCheckBox);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.CheckBox.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.CheckBox.docx");

        ArrayPtr<SharedPtr<StructuredDocumentTag>> sdts =
            doc->GetChildNodes(NodeType::StructuredDocumentTag, true)->LINQ_OfType<SharedPtr<StructuredDocumentTag>>()->LINQ_ToArray();

        ASPOSE_ASSERT_EQ(true, sdts[0]->get_Checked());
        ASSERT_EQ(sdts[0]->get_XmlMapping()->get_StoreItemId(), String::Empty);
    }

    void Date()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.CalendarType
        //ExFor:StructuredDocumentTag.DateDisplayFormat
        //ExFor:StructuredDocumentTag.DateDisplayLocale
        //ExFor:StructuredDocumentTag.DateStorageFormat
        //ExFor:StructuredDocumentTag.FullDate
        //ExSummary:Shows how to prompt the user to enter a date with a structured document tag.
        auto doc = MakeObject<Document>();

        // Insert a structured document tag that prompts the user to enter a date.
        // In Microsoft Word, this element is known as a "Date picker content control".
        // When we click on the arrow on the right end of this tag in Microsoft Word,
        // we will see a pop up in the form of a clickable calendar.
        // We can use that popup to select a date that the tag will display.
        auto sdtDate = MakeObject<StructuredDocumentTag>(doc, SdtType::Date, MarkupLevel::Inline);

        // Display the date, according to the Saudi Arabian Arabic locale.
        sdtDate->set_DateDisplayLocale(System::Globalization::CultureInfo::GetCultureInfo(u"ar-SA")->get_LCID());

        // Set the format with which to display the date.
        sdtDate->set_DateDisplayFormat(u"dd MMMM, yyyy");
        sdtDate->set_DateStorageFormat(SdtDateStorageFormat::DateTime);

        // Display the date according to the Hijri calendar.
        sdtDate->set_CalendarType(SdtCalendarType::Hijri);

        // Before the user chooses a date in Microsoft Word, the tag will display the text "Click here to enter a date.".
        // According to the tag's calendar, set the "FullDate" property to get the tag to display a default date.
        sdtDate->set_FullDate(System::DateTime(1440, 10, 20));

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
        //ExSummary:Shows how to create a structured document tag in a plain text box and modify its appearance.
        auto doc = MakeObject<Document>();

        // Create a structured document tag that will contain plain text.
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // Set the title and color of the frame that appears when you mouse over the structured document tag in Microsoft Word.
        tag->set_Title(u"My plain text");
        tag->set_Color(System::Drawing::Color::get_Magenta());

        // Set a tag for this structured document tag, which is obtainable
        // as an XML element named "tag", with the string below in its "@val" attribute.
        tag->set_Tag(u"MyPlainTextSDT");

        // Every structured document tag has a random unique ID.
        ASSERT_GT(tag->get_Id(), 0);

        // Set the font for the text inside the structured document tag.
        tag->get_ContentsFont()->set_Name(u"Arial");

        // Set the font for the text at the end of the structured document tag.
        // Any text that we type in the document body after moving out of the tag with arrow keys will use this font.
        tag->get_EndCharacterFont()->set_Name(u"Arial Black");

        // By default, this is false and pressing enter while inside a structured document tag does nothing.
        // When set to true, our structured document tag can have multiple lines.

        // Set the "Multiline" property to "false" to only allow the contents
        // of this structured document tag to span a single line.
        // Set the "Multiline" property to "true" to allow the tag to contain multiple lines of content.
        tag->set_Multiline(true);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertNode(tag);

        // Insert a clone of our structured document tag in a new paragraph.
        auto tagClone = System::DynamicCast<StructuredDocumentTag>(tag->Clone(true));
        builder->InsertParagraph();
        builder->InsertNode(tagClone);

        // Use the "RemoveSelfOnly" method to remove a structured document tag, while keeping its contents in the document.
        tagClone->RemoveSelfOnly();

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.PlainText.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.PlainText.docx");
        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));

        ASSERT_EQ(u"My plain text", tag->get_Title());
        ASSERT_EQ(System::Drawing::Color::get_Magenta().ToArgb(), tag->get_Color().ToArgb());
        ASSERT_EQ(u"MyPlainTextSDT", tag->get_Tag());
        ASSERT_GT(tag->get_Id(), 0);
        ASSERT_EQ(u"Arial", tag->get_ContentsFont()->get_Name());
        ASSERT_EQ(u"Arial Black", tag->get_EndCharacterFont()->get_Name());
        ASSERT_TRUE(tag->get_Multiline());
    }

    void IsTemporary(bool isTemporary)
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsTemporary
        //ExSummary:Shows how to make single-use controls.
        auto doc = MakeObject<Document>();

        // Insert a plain text structured document tag,
        // which will act as a plain text form that the user may enter text into.
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // Set the "IsTemporary" property to "true" to make the structured document tag disappear and
        // assimilate its contents into the document after the user edits it once in Microsoft Word.
        // Set the "IsTemporary" property to "false" to allow the user to edit the contents
        // of the structured document tag any number of times.
        tag->set_IsTemporary(isTemporary);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Please enter text: ");
        builder->InsertNode(tag);

        // Insert another structured document tag in the form of a check box and set its default state to "checked".
        tag = MakeObject<StructuredDocumentTag>(doc, SdtType::Checkbox, MarkupLevel::Inline);
        tag->set_Checked(true);

        // Set the "IsTemporary" property to "true" to make the check box become a symbol
        // once the user clicks on it in Microsoft Word.
        // Set the "IsTemporary" property to "false" to allow the user to click on the check box any number of times.
        tag->set_IsTemporary(isTemporary);

        builder->Write(u"\nPlease click the check box: ");
        builder->InsertNode(tag);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.IsTemporary.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.IsTemporary.docx");

        ASSERT_EQ(2, doc->GetChildNodes(NodeType::StructuredDocumentTag, true)
                         ->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(static_cast<std::function<bool(SharedPtr<Node> sdt)>>(
                             [&isTemporary](SharedPtr<Node> sdt) -> bool
                             { return (System::DynamicCast<StructuredDocumentTag>(sdt))->get_IsTemporary() == isTemporary; }))));
    }

    void PlaceholderBuildingBlock(bool isShowingPlaceholderText)
    {
        //ExStart
        //ExFor:StructuredDocumentTag.IsShowingPlaceholderText
        //ExFor:StructuredDocumentTag.Placeholder
        //ExFor:StructuredDocumentTag.PlaceholderName
        //ExSummary:Shows how to use a building block's contents as a custom placeholder text for a structured document tag.
        auto doc = MakeObject<Document>();

        // Insert a plain text structured document tag of the "PlainText" type, which will function as a text box.
        // The contents that it will display by default are a "Click here to enter text." prompt.
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // We can get the tag to display the contents of a building block instead of the default text.
        // First, add a building block with contents to the glossary document.
        SharedPtr<GlossaryDocument> glossaryDoc = doc->get_GlossaryDocument();

        auto substituteBlock = MakeObject<BuildingBlock>(glossaryDoc);
        substituteBlock->set_Name(u"Custom Placeholder");
        substituteBlock->AppendChild(MakeObject<Section>(glossaryDoc));
        substituteBlock->get_FirstSection()->AppendChild(MakeObject<Body>(glossaryDoc));
        substituteBlock->get_FirstSection()->get_Body()->AppendParagraph(u"Custom placeholder text.");

        glossaryDoc->AppendChild(substituteBlock);

        // Then, use the structured document tag's "PlaceholderName" property to reference that building block by name.
        tag->set_PlaceholderName(u"Custom Placeholder");

        // If "PlaceholderName" refers to an existing block in the parent document's glossary document,
        // we will be able to verify the building block via the "Placeholder" property.
        ASPOSE_ASSERT_EQ(substituteBlock, tag->get_Placeholder());

        // Set the "IsShowingPlaceholderText" property to "true" to treat the
        // structured document tag's current contents as placeholder text.
        // This means that clicking on the text box in Microsoft Word will immediately highlight all the tag's contents.
        // Set the "IsShowingPlaceholderText" property to "false" to get the
        // structured document tag to treat its contents as text that a user has already entered.
        // Clicking on this text in Microsoft Word will place the blinking cursor at the clicked location.
        tag->set_IsShowingPlaceholderText(isShowingPlaceholderText);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertNode(tag);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.PlaceholderBuildingBlock.docx");
        tag = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        substituteBlock = System::DynamicCast<BuildingBlock>(doc->get_GlossaryDocument()->GetChild(NodeType::BuildingBlock, 0, true));

        ASSERT_EQ(u"Custom Placeholder", substituteBlock->get_Name());
        ASPOSE_ASSERT_EQ(isShowingPlaceholderText, tag->get_IsShowingPlaceholderText());
        ASPOSE_ASSERT_EQ(substituteBlock, tag->get_Placeholder());
        ASSERT_EQ(substituteBlock->get_Name(), tag->get_PlaceholderName());
    }

    void Lock()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.LockContentControl
        //ExFor:StructuredDocumentTag.LockContents
        //ExSummary:Shows how to apply editing restrictions to structured document tags.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a plain text structured document tag, which acts as a text box that prompts the user to fill it in.
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // Set the "LockContents" property to "true" to prohibit the user from editing this text box's contents.
        tag->set_LockContents(true);
        builder->Write(u"The contents of this structured document tag cannot be edited: ");
        builder->InsertNode(tag);

        tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Inline);

        // Set the "LockContentControl" property to "true" to prohibit the user from
        // deleting this structured document tag manually in Microsoft Word.
        tag->set_LockContentControl(true);

        builder->InsertParagraph();
        builder->Write(u"This structured document tag cannot be deleted but its contents can be edited: ");
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
        //ExSummary:Shows how to work with drop down-list structured document tags.
        auto doc = MakeObject<Document>();
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::DropDownList, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        // A drop-down list structured document tag is a form that allows the user to
        // select an option from a list by left-clicking and opening the form in Microsoft Word.
        // The "ListItems" property contains all list items, and each list item is an "SdtListItem".
        SharedPtr<SdtListItemCollection> listItems = tag->get_ListItems();
        listItems->Add(MakeObject<SdtListItem>(u"Value 1"));

        ASSERT_EQ(listItems->idx_get(0)->get_DisplayText(), listItems->idx_get(0)->get_Value());

        // Add 3 more list items. Initialize these items using a different constructor to the first item
        // to display strings that are different from their values.
        listItems->Add(MakeObject<SdtListItem>(u"Item 2", u"Value 2"));
        listItems->Add(MakeObject<SdtListItem>(u"Item 3", u"Value 3"));
        listItems->Add(MakeObject<SdtListItem>(u"Item 4", u"Value 4"));

        ASSERT_EQ(4, listItems->get_Count());

        // The drop-down list is displaying the first item. Assign a different list item to the "SelectedValue" to display it.
        listItems->set_SelectedValue(listItems->idx_get(3));

        ASSERT_EQ(u"Value 4", listItems->get_SelectedValue()->get_Value());

        // Enumerate over the collection and print each element.
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

        // Remove the last list item.
        listItems->RemoveAt(3);

        ASSERT_EQ(3, listItems->get_Count());

        // Since our drop-down control is set to display the removed item by default, give it an item to display which exists.
        listItems->set_SelectedValue(listItems->idx_get(1));

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.ListItemCollection.docx");

        // Use the "Clear" method to empty the entire drop-down item collection at once.
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
        //ExSummary:Shows how to create a structured document tag with custom XML data.
        auto doc = MakeObject<Document>();

        // Construct an XML part that contains data and add it to the document's collection.
        // If we enable the "Developer" tab in Microsoft Word,
        // we can find elements from this collection in the "XML Mapping Pane", along with a few default elements.
        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Hello world!</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);

        ASPOSE_ASSERT_EQ(System::Text::Encoding::get_ASCII()->GetBytes(xmlPartContent), xmlPart->get_Data());
        ASSERT_EQ(xmlPartId, xmlPart->get_Id());

        // Below are two ways to refer to XML parts.
        // 1 -  By an index in the custom XML part collection:
        ASPOSE_ASSERT_EQ(xmlPart, doc->get_CustomXmlParts()->idx_get(0));

        // 2 -  By GUID:
        ASPOSE_ASSERT_EQ(xmlPart, doc->get_CustomXmlParts()->GetById(xmlPartId));

        // Add an XML schema association.
        xmlPart->get_Schemas()->Add(u"http://www.w3.org/2001/XMLSchema");

        // Clone a part, and then insert it into the collection.
        SharedPtr<CustomXmlPart> xmlPartClone = xmlPart->Clone();
        xmlPartClone->set_Id(System::Guid::NewGuid().ToString(u"B"));
        doc->get_CustomXmlParts()->Add(xmlPartClone);

        ASSERT_EQ(2, doc->get_CustomXmlParts()->get_Count());

        // Iterate through the collection and print the contents of each part.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<CustomXmlPart>>> enumerator = doc->get_CustomXmlParts()->GetEnumerator();
            int index = 0;
            while (enumerator->MoveNext())
            {
                std::cout << "XML part index " << index << ", ID: " << enumerator->get_Current()->get_Id() << std::endl;
                std::cout << "\tContent: " << System::Text::Encoding::get_UTF8()->GetString(enumerator->get_Current()->get_Data()) << std::endl;
                index++;
            }
        }

        // Use the "RemoveAt" method to remove the cloned part by index.
        doc->get_CustomXmlParts()->RemoveAt(1);

        ASSERT_EQ(1, doc->get_CustomXmlParts()->get_Count());

        // Clone the XML parts collection, and then use the "Clear" method to remove all its elements at once.
        SharedPtr<CustomXmlPartCollection> customXmlParts = doc->get_CustomXmlParts()->Clone();
        customXmlParts->Clear();

        // Create a structured document tag that will display our part's contents and insert it into the document body.
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
        //ExSummary:Shows how to set XML mappings for custom XML parts.
        auto doc = MakeObject<Document>();

        // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Text element #1</text><text>Text element #2</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);

        ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));

        // Create a structured document tag that will display the contents of our CustomXmlPart.
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);

        // Set a mapping for our structured document tag. This mapping will instruct
        // our structured document tag to display a portion of the XML part's text contents that the XPath points to.
        // In this case, it will be contents of the the second "<text>" element of the first "<root>" element: "Text element #2".
        tag->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[2]", u"xmlns:ns='http://www.w3.org/2001/XMLSchema'");

        ASSERT_TRUE(tag->get_XmlMapping()->get_IsMapped());
        ASPOSE_ASSERT_EQ(xmlPart, tag->get_XmlMapping()->get_CustomXmlPart());
        ASSERT_EQ(u"/root[1]/text[2]", tag->get_XmlMapping()->get_XPath());
        ASSERT_EQ(u"xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag->get_XmlMapping()->get_PrefixMappings());

        // Add the structured document tag to the document to display the content from our custom part.
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

    void StructuredDocumentTagRangeStartXmlMapping()
    {
        //ExStart
        //ExFor:StructuredDocumentTagRangeStart.XmlMapping
        //ExSummary:Shows how to set XML mappings for the range start of a structured document tag.
        auto doc = MakeObject<Document>(MyDir + u"Multi-section structured document tags.docx");

        // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Text element #1</text><text>Text element #2</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);

        ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));

        // Create a structured document tag that will display the contents of our CustomXmlPart in the document.
        auto sdtRangeStart = System::DynamicCast<StructuredDocumentTagRangeStart>(doc->GetChild(NodeType::StructuredDocumentTagRangeStart, 0, true));

        // If we set a mapping for our structured document tag,
        // it will only display a portion of the CustomXmlPart that the XPath points to.
        // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
        sdtRangeStart->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[2]", nullptr);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
        xmlPart = doc->get_CustomXmlParts()->idx_get(0);

        System::Guid temp;
        ASSERT_TRUE(System::Guid::TryParse(xmlPart->get_Id(), temp));
        ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));

        sdtRangeStart = System::DynamicCast<StructuredDocumentTagRangeStart>(doc->GetChild(NodeType::StructuredDocumentTagRangeStart, 0, true));
        ASSERT_EQ(u"/root[1]/text[2]", sdtRangeStart->get_XmlMapping()->get_XPath());
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
        auto doc = MakeObject<Document>();

        String xmlPartId = System::Guid::NewGuid().ToString(u"B");
        String xmlPartContent = u"<root><text>Hello, World!</text></root>";
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);

        // Add an XML schema association.
        xmlPart->get_Schemas()->Add(u"http://www.w3.org/2001/XMLSchema");

        // Clone the custom XML part's XML schema association collection,
        // and then add a couple of new schemas to the clone.
        SharedPtr<CustomXmlSchemaCollection> schemas = xmlPart->get_Schemas()->Clone();
        schemas->Add(u"http://www.w3.org/2001/XMLSchema-instance");
        schemas->Add(u"http://schemas.microsoft.com/office/2006/metadata/contentType");

        ASSERT_EQ(3, schemas->get_Count());
        ASSERT_EQ(2, schemas->IndexOf(u"http://schemas.microsoft.com/office/2006/metadata/contentType"));

        // Enumerate the schemas and print each element.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<String>> enumerator = schemas->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << enumerator->get_Current() << std::endl;
            }
        }

        // Below are three ways of removing schemas from the collection.
        // 1 -  Remove a schema by index:
        schemas->RemoveAt(2);

        // 2 -  Remove a schema by value:
        schemas->Remove(u"http://www.w3.org/2001/XMLSchema");

        // 3 -  Use the "Clear" method to empty the collection at once.
        schemas->Clear();

        ASSERT_EQ(0, schemas->get_Count());
        //ExEnd
    }

    void CustomXmlPartStoreItemIdReadOnly()
    {
        //ExStart
        //ExFor:XmlMapping.StoreItemId
        //ExSummary:Shows how to get the custom XML data identifier of an XML part.
        auto doc = MakeObject<Document>(MyDir + u"Custom XML part in structured document tag.docx");

        // Structured document tags have IDs in the form of GUIDs.
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

        builder->InsertNode(sdtCheckBox);

        doc = DocumentHelper::SaveOpen(doc);

        auto sdt = System::DynamicCast<StructuredDocumentTag>(doc->GetChild(NodeType::StructuredDocumentTag, 0, true));
        std::cout << (String(u"The Id of your custom xml part is: ") + sdt->get_XmlMapping()->get_StoreItemId()) << std::endl;
    }

    void ClearTextFromStructuredDocumentTags()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.Clear
        //ExSummary:Shows how to delete contents of structured document tag elements.
        auto doc = MakeObject<Document>();

        // Create a plain text structured document tag, and then append it to the document.
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        // This structured document tag, which is in the form of a text box, already displays placeholder text.
        ASSERT_EQ(u"Click here to enter text.", tag->GetText().Trim());
        ASSERT_TRUE(tag->get_IsShowingPlaceholderText());

        // Create a building block with text contents.
        SharedPtr<GlossaryDocument> glossaryDoc = doc->get_GlossaryDocument();
        auto substituteBlock = MakeObject<BuildingBlock>(glossaryDoc);
        substituteBlock->set_Name(u"My placeholder");
        substituteBlock->AppendChild(MakeObject<Section>(glossaryDoc));
        substituteBlock->get_FirstSection()->EnsureMinimum();
        substituteBlock->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(glossaryDoc, u"Custom placeholder text."));
        glossaryDoc->AppendChild(substituteBlock);

        // Set the structured document tag's "PlaceholderName" property to our building block's name to get
        // the structured document tag to display the contents of the building block in place of the original default text.
        tag->set_PlaceholderName(u"My placeholder");

        ASSERT_EQ(u"Custom placeholder text.", tag->GetText().Trim());
        ASSERT_TRUE(tag->get_IsShowingPlaceholderText());

        // Edit the text of the structured document tag and hide the placeholder text.
        auto run = System::DynamicCast<Run>(tag->GetChild(NodeType::Run, 0, true));
        run->set_Text(u"New text.");
        tag->set_IsShowingPlaceholderText(false);

        ASSERT_EQ(u"New text.", tag->GetText().Trim());

        // Use the "Clear" method to clear this structured document tag's contents and display the placeholder again.
        tag->Clear();

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
        ASSERT_THROW(static_cast<std::function<String()>>([&plainTextSdt]() -> String { return plainTextSdt->get_BuildingBlockGallery(); })(),
                     System::InvalidOperationException);
    }

    void BuildingBlockCategories()
    {
        //ExStart
        //ExFor:StructuredDocumentTag.BuildingBlockCategory
        //ExFor:StructuredDocumentTag.BuildingBlockGallery
        //ExSummary:Shows how to insert a structured document tag as a building block, and set its category and gallery.
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
        //ExSummary:Shows how to update structured document tags while saving a document to PDF.
        auto doc = MakeObject<Document>();

        // Insert a drop-down list structured document tag.
        auto tag = MakeObject<StructuredDocumentTag>(doc, SdtType::DropDownList, MarkupLevel::Block);
        tag->get_ListItems()->Add(MakeObject<SdtListItem>(u"Value 1"));
        tag->get_ListItems()->Add(MakeObject<SdtListItem>(u"Value 2"));
        tag->get_ListItems()->Add(MakeObject<SdtListItem>(u"Value 3"));

        // The drop-down list currently displays "Choose an item" as the default text.
        // Set the "SelectedValue" property to one of the list items to get the tag to
        // display that list item's value instead of the default text.
        tag->get_ListItems()->set_SelectedValue(tag->get_ListItems()->idx_get(1));

        doc->get_FirstSection()->get_Body()->AppendChild(tag);

        // Create a "PdfSaveOptions" object to pass to the document's "Save" method
        // to modify how that method saves the document to .PDF.
        auto options = MakeObject<PdfSaveOptions>();

        // Set the "UpdateSdtContent" property to "false" not to update the structured document tags
        // while saving the document to PDF. They will display their default values as they were at the time of construction.
        // Set the "UpdateSdtContent" property to "true" to make sure the tags display updated values in the PDF.
        options->set_UpdateSdtContent(updateSdtContent);

        doc->Save(ArtifactsDir + u"StructuredDocumentTag.UpdateSdtContent.pdf", options);
        //ExEnd
    }

    void FillTableUsingRepeatingSectionItem()
    {
        //ExStart
        //ExFor:SdtType
        //ExSummary:Shows how to fill a table with data from in an XML part.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(
            u"Books", String(u"<books>") + u"<book>" + u"<title>Everyday Italian</title>" + u"<author>Giada De Laurentiis</author>" + u"</book>" + u"<book>" +
                          u"<title>The C Programming Language</title>" + u"<author>Brian W. Kernighan, Dennis M. Ritchie</author>" + u"</book>" + u"<book>" +
                          u"<title>Learning XML</title>" + u"<author>Erik T. Ray</author>" + u"</book>" + u"</books>");

        // Create headers for data from the XML content.
        SharedPtr<Table> table = builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Title");
        builder->InsertCell();
        builder->Write(u"Author");
        builder->EndRow();
        builder->EndTable();

        // Create a table with a repeating section inside.
        auto repeatingSectionSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::RepeatingSection, MarkupLevel::Row);
        repeatingSectionSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book", String::Empty);
        table->AppendChild(repeatingSectionSdt);

        // Add repeating section item inside the repeating section and mark it as a row.
        // This table will have a row for each element that we can find in the XML document
        // using the "/books[1]/book" XPath, of which there are three.
        auto repeatingSectionItemSdt = MakeObject<StructuredDocumentTag>(doc, SdtType::RepeatingSectionItem, MarkupLevel::Row);
        repeatingSectionSdt->AppendChild(repeatingSectionItemSdt);

        auto row = MakeObject<Row>(doc);
        repeatingSectionItemSdt->AppendChild(row);

        // Map XML data with created table cells for the title and author of each book.
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
                      u"The C Programming Language\u0007Brian W. Kernighan, Dennis M. Ritchie\u0007\u0007" + u"Learning XML\u0007Erik T. Ray\u0007\u0007",
                  doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->GetText().Trim());
    }

    void CustomXmlPart_()
    {
        String xmlString = String(u"<?xml version=\"1.0\"?>") + u"<Company>" + u"<Employee id=\"1\">" + u"<FirstName>John</FirstName>" +
                           u"<LastName>Doe</LastName>" + u"</Employee>" + u"<Employee id=\"2\">" + u"<FirstName>Jane</FirstName>" +
                           u"<LastName>Doe</LastName>" + u"</Employee>" + u"</Company>";

        auto doc = MakeObject<Document>();

        // Insert the full XML document as a custom document part.
        // We can find the mapping for this part in Microsoft Word via "Developer" -> "XML Mapping Pane", if it is enabled.
        SharedPtr<CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(System::Guid::NewGuid().ToString(u"B"), xmlString);

        // Create a structured document tag, which will use an XPath to refer to a single element from the XML.
        auto sdt = MakeObject<StructuredDocumentTag>(doc, SdtType::PlainText, MarkupLevel::Block);
        sdt->get_XmlMapping()->SetMapping(xmlPart, u"Company//Employee[@id='2']/FirstName", u"");

        // Add the StructuredDocumentTag to the document to display the element in the text.
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
        //ExSummary:Shows how to get the properties of multi-section structured document tags.
        auto doc = MakeObject<Document>(MyDir + u"Multi-section structured document tags.docx");

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

    void SdtChildNodes()
    {
        //ExStart
        //ExFor:StructuredDocumentTagRangeStart.ChildNodes
        //ExFor:StructuredDocumentTagRangeStart.GetChildNodes(NodeType, bool)
        //ExSummary:Shows how to get child nodes of StructuredDocumentTagRangeStart.
        auto doc = MakeObject<Document>(MyDir + u"Multi-section structured document tags.docx");
        auto tag =
            System::DynamicCast_noexcept<StructuredDocumentTagRangeStart>(doc->GetChildNodes(NodeType::StructuredDocumentTagRangeStart, true)->idx_get(0));

        std::cout << "StructuredDocumentTagRangeStart values:" << std::endl;
        std::cout << "\t|Child nodes count: " << tag->get_ChildNodes()->get_Count() << "\n" << std::endl;

        for (const auto& node : System::IterateOver(tag->get_ChildNodes()))
        {
            std::cout << String::Format(u"\t|Child node type: {0}", node->get_NodeType()) << std::endl;
        }

        for (const auto& node : System::IterateOver(tag->GetChildNodes(NodeType::Run, true)))
        {
            std::cout << "\t|Child node text: " << node->GetText() << std::endl;
        }
        //ExEnd
    }
};

} // namespace ApiExamples
