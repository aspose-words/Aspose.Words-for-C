// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExStructuredDocumentTag.h"

#include <testing/test_predicates.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/guid.h>
#include <system/globalization/culture_info.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/default.h>
#include <system/date_time.h>
#include <system/convert.h>
#include <system/collections/list.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/XmlMapping.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTagCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtType.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtListItemCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtListItem.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtDateStorageFormat.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtCalendarType.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/SdtAppearance.h>
#include <Aspose.Words.Cpp/Model/Markup/Sdt/IStructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Model/Markup/MarkupLevel.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlSchemaCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPartCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomXmlPart.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlock.h>

#include "DocumentHelper.h"


using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(632790078u, ::Aspose::Words::ApiExamples::ExStructuredDocumentTag, ThisTypeBaseTypesInfo);

System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeStart> ExStructuredDocumentTag::InsertStructuredDocumentTagRanges(System::SharedPtr<Aspose::Words::Document> doc)
{
    auto rangeStart = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc, Aspose::Words::Markup::SdtType::PlainText);
    auto rangeEnd = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTagRangeEnd>(doc, rangeStart->get_Id());
    
    doc->get_FirstSection()->get_Body()->InsertBefore<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeStart>>(rangeStart, doc->get_FirstSection()->get_Body()->get_FirstParagraph());
    doc->get_LastSection()->get_Body()->InsertAfter<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeEnd>>(rangeEnd, doc->get_FirstSection()->get_Body()->get_FirstParagraph());
    
    return rangeStart;
}


namespace gtest_test
{

class ExStructuredDocumentTag : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExStructuredDocumentTag> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExStructuredDocumentTag>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExStructuredDocumentTag> ExStructuredDocumentTag::s_instance;

} // namespace gtest_test

void ExStructuredDocumentTag::RepeatingSection()
{
    //ExStart
    //ExFor:StructuredDocumentTag.SdtType
    //ExFor:IStructuredDocumentTag.SdtType
    //ExSummary:Shows how to get the type of a structured document tag.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags.docx");
    
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>> tags = doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTag, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> >()->LINQ_ToList();
    
    ASSERT_EQ(Aspose::Words::Markup::SdtType::RepeatingSection, tags->idx_get(0)->get_SdtType());
    ASSERT_EQ(Aspose::Words::Markup::SdtType::RepeatingSectionItem, tags->idx_get(1)->get_SdtType());
    ASSERT_EQ(Aspose::Words::Markup::SdtType::RichText, tags->idx_get(2)->get_SdtType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, RepeatingSection)
{
    s_instance->RepeatingSection();
}

} // namespace gtest_test

void ExStructuredDocumentTag::FlatOpcContent()
{
    //ExStart
    //ExFor:StructuredDocumentTag.WordOpenXML
    //ExFor:IStructuredDocumentTag.WordOpenXML
    //ExSummary:Shows how to get XML contained within the node in the FlatOpc format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags.docx");
    
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>> tags = doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTag, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> >()->LINQ_ToList();
    
    ASSERT_TRUE(tags->idx_get(0)->get_WordOpenXML().Contains(u"<pkg:part pkg:name=\"/docProps/app.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\">"));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, FlatOpcContent)
{
    s_instance->FlatOpcContent();
}

} // namespace gtest_test

void ExStructuredDocumentTag::ApplyStyle()
{
    //ExStart
    //ExFor:StructuredDocumentTag
    //ExFor:StructuredDocumentTag.NodeType
    //ExFor:StructuredDocumentTag.Style
    //ExFor:StructuredDocumentTag.StyleName
    //ExFor:StructuredDocumentTag.WordOpenXMLMinimal
    //ExFor:MarkupLevel
    //ExFor:SdtType
    //ExSummary:Shows how to work with styles for content control elements.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two ways to apply a style from the document to a structured document tag.
    // 1 -  Apply a style object from the document's style collection:
    System::SharedPtr<Aspose::Words::Style> quoteStyle = doc->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Quote);
    auto sdtPlainText = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Inline);
    sdtPlainText->set_Style(quoteStyle);
    
    // 2 -  Reference a style in the document by name:
    auto sdtRichText = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::RichText, Aspose::Words::Markup::MarkupLevel::Inline);
    sdtRichText->set_StyleName(u"Quote");
    
    builder->InsertNode(sdtPlainText);
    builder->InsertNode(sdtRichText);
    
    ASSERT_EQ(Aspose::Words::NodeType::StructuredDocumentTag, sdtPlainText->get_NodeType());
    
    System::SharedPtr<Aspose::Words::NodeCollection> tags = doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTag, true);
    
    for (auto&& node : System::IterateOver(tags))
    {
        auto sdt = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(node);
        
        std::cout << sdt->get_WordOpenXMLMinimal() << std::endl;
        
        ASSERT_EQ(Aspose::Words::StyleIdentifier::Quote, sdt->get_Style()->get_StyleIdentifier());
        ASSERT_EQ(u"Quote", sdt->get_StyleName());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, ApplyStyle)
{
    s_instance->ApplyStyle();
}

} // namespace gtest_test

void ExStructuredDocumentTag::CheckBox()
{
    //ExStart
    //ExFor:StructuredDocumentTag.#ctor(DocumentBase, SdtType, MarkupLevel)
    //ExFor:StructuredDocumentTag.Checked
    //ExFor:StructuredDocumentTag.SetCheckedSymbol(Int32, String)
    //ExFor:StructuredDocumentTag.SetUncheckedSymbol(Int32, String)
    //ExSummary:Show how to create a structured document tag in the form of a check box.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto sdtCheckBox = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::Checkbox, Aspose::Words::Markup::MarkupLevel::Inline);
    sdtCheckBox->set_Checked(true);
    
    // We can set the symbols used to represent the checked/unchecked state of a checkbox content control.
    sdtCheckBox->SetCheckedSymbol(0x00A9, u"Times New Roman");
    sdtCheckBox->SetUncheckedSymbol(0x00AE, u"Times New Roman");
    
    builder->InsertNode(sdtCheckBox);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.CheckBox.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.CheckBox.docx");
    
    System::ArrayPtr<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>> tags = doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTag, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> >()->LINQ_ToArray();
    
    ASPOSE_ASSERT_EQ(true, tags[0]->get_Checked());
    ASSERT_EQ(System::String::Empty, tags[0]->get_XmlMapping()->get_StoreItemId());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, CheckBox)
{
    s_instance->CheckBox();
}

} // namespace gtest_test

void ExStructuredDocumentTag::Date()
{
    //ExStart
    //ExFor:StructuredDocumentTag.CalendarType
    //ExFor:StructuredDocumentTag.DateDisplayFormat
    //ExFor:StructuredDocumentTag.DateDisplayLocale
    //ExFor:StructuredDocumentTag.DateStorageFormat
    //ExFor:StructuredDocumentTag.FullDate
    //ExFor:SdtCalendarType
    //ExFor:SdtDateStorageFormat
    //ExSummary:Shows how to prompt the user to enter a date with a structured document tag.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert a structured document tag that prompts the user to enter a date.
    // In Microsoft Word, this element is known as a "Date picker content control".
    // When we click on the arrow on the right end of this tag in Microsoft Word,
    // we will see a pop up in the form of a clickable calendar.
    // We can use that popup to select a date that the tag will display.
    auto sdtDate = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::Date, Aspose::Words::Markup::MarkupLevel::Inline);
    
    // Display the date, according to the Saudi Arabian Arabic locale.
    sdtDate->set_DateDisplayLocale(System::Globalization::CultureInfo::GetCultureInfo(u"ar-SA")->get_LCID());
    
    // Set the format with which to display the date.
    sdtDate->set_DateDisplayFormat(u"dd MMMM, yyyy");
    sdtDate->set_DateStorageFormat(Aspose::Words::Markup::SdtDateStorageFormat::DateTime);
    
    // Display the date according to the Hijri calendar.
    sdtDate->set_CalendarType(Aspose::Words::Markup::SdtCalendarType::Hijri);
    
    // Before the user chooses a date in Microsoft Word, the tag will display the text "Click here to enter a date.".
    // According to the tag's calendar, set the "FullDate" property to get the tag to display a default date.
    sdtDate->set_FullDate(System::DateTime(1440, 10, 20));
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertNode(sdtDate);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.Date.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, SkipMono_Date)
{
    RecordProperty("category", "SkipMono");
    
    s_instance->Date();
}

} // namespace gtest_test

void ExStructuredDocumentTag::PlainText()
{
    //ExStart
    //ExFor:StructuredDocumentTag.Color
    //ExFor:StructuredDocumentTag.ContentsFont
    //ExFor:StructuredDocumentTag.EndCharacterFont
    //ExFor:StructuredDocumentTag.Id
    //ExFor:StructuredDocumentTag.Level
    //ExFor:StructuredDocumentTag.Multiline
    //ExFor:IStructuredDocumentTag.Tag
    //ExFor:StructuredDocumentTag.Tag
    //ExFor:StructuredDocumentTag.Title
    //ExFor:StructuredDocumentTag.RemoveSelfOnly
    //ExFor:StructuredDocumentTag.Appearance
    //ExSummary:Shows how to create a structured document tag in a plain text box and modify its appearance.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a structured document tag that will contain plain text.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Inline);
    
    // Set the title and color of the frame that appears when you mouse over the structured document tag in Microsoft Word.
    tag->set_Title(u"My plain text");
    tag->set_Color(System::Drawing::Color::get_Magenta());
    
    // Set a tag for this structured document tag, which is obtainable
    // as an XML element named "tag", with the string below in its "@val" attribute.
    tag->set_Tag(u"MyPlainTextSDT");
    
    // Every structured document tag has a random unique ID.
    ASSERT_TRUE(tag->get_Id() > 0);
    
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
    
    // Set the "Appearance" property to "SdtAppearance.Tags" to show tags around content.
    // By default structured document tag shows as BoundingBox. 
    tag->set_Appearance(Aspose::Words::Markup::SdtAppearance::Tags);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertNode(tag);
    
    // Insert a clone of our structured document tag in a new paragraph.
    auto tagClone = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(System::ExplicitCast<Aspose::Words::Node>(tag)->Clone(true));
    builder->InsertParagraph();
    builder->InsertNode(tagClone);
    
    // Use the "RemoveSelfOnly" method to remove a structured document tag, while keeping its contents in the document.
    tagClone->RemoveSelfOnly();
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.PlainText.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.PlainText.docx");
    tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    
    ASSERT_EQ(u"My plain text", tag->get_Title());
    ASSERT_EQ(System::Drawing::Color::get_Magenta().ToArgb(), tag->get_Color().ToArgb());
    ASSERT_EQ(u"MyPlainTextSDT", tag->get_Tag());
    ASSERT_TRUE(tag->get_Id() > 0);
    ASSERT_EQ(u"Arial", tag->get_ContentsFont()->get_Name());
    ASSERT_EQ(u"Arial Black", tag->get_EndCharacterFont()->get_Name());
    ASSERT_TRUE(tag->get_Multiline());
    ASSERT_EQ(Aspose::Words::Markup::SdtAppearance::Tags, tag->get_Appearance());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, PlainText)
{
    s_instance->PlainText();
}

} // namespace gtest_test

void ExStructuredDocumentTag::IsTemporary(bool isTemporary)
{
    //ExStart
    //ExFor:StructuredDocumentTag.IsTemporary
    //ExSummary:Shows how to make single-use controls.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert a plain text structured document tag,
    // which will act as a plain text form that the user may enter text into.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Inline);
    
    // Set the "IsTemporary" property to "true" to make the structured document tag disappear and
    // assimilate its contents into the document after the user edits it once in Microsoft Word.
    // Set the "IsTemporary" property to "false" to allow the user to edit the contents
    // of the structured document tag any number of times.
    tag->set_IsTemporary(isTemporary);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"Please enter text: ");
    builder->InsertNode(tag);
    
    // Insert another structured document tag in the form of a check box and set its default state to "checked".
    tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::Checkbox, Aspose::Words::Markup::MarkupLevel::Inline);
    tag->set_Checked(true);
    
    // Set the "IsTemporary" property to "true" to make the check box become a symbol
    // once the user clicks on it in Microsoft Word.
    // Set the "IsTemporary" property to "false" to allow the user to click on the check box any number of times.
    tag->set_IsTemporary(isTemporary);
    
    builder->Write(u"\nPlease click the check box: ");
    builder->InsertNode(tag);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.IsTemporary.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.IsTemporary.docx");
    
    ASSERT_EQ(2, doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTag, true)->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> sdt)>>([&isTemporary](System::SharedPtr<Aspose::Words::Node> sdt) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(sdt))->get_IsTemporary() == isTemporary;
    }))));
}

namespace gtest_test
{

using ExStructuredDocumentTag_IsTemporary_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExStructuredDocumentTag::IsTemporary)>::type;

struct ExStructuredDocumentTag_IsTemporary : public ExStructuredDocumentTag, public Aspose::Words::ApiExamples::ExStructuredDocumentTag, public ::testing::WithParamInterface<ExStructuredDocumentTag_IsTemporary_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExStructuredDocumentTag_IsTemporary, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IsTemporary(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExStructuredDocumentTag_IsTemporary, ::testing::ValuesIn(ExStructuredDocumentTag_IsTemporary::TestCases()));

} // namespace gtest_test

void ExStructuredDocumentTag::PlaceholderBuildingBlock(bool isShowingPlaceholderText)
{
    //ExStart
    //ExFor:StructuredDocumentTag.IsShowingPlaceholderText
    //ExFor:IStructuredDocumentTag.IsShowingPlaceholderText
    //ExFor:StructuredDocumentTag.Placeholder
    //ExFor:StructuredDocumentTag.PlaceholderName
    //ExFor:IStructuredDocumentTag.Placeholder
    //ExFor:IStructuredDocumentTag.PlaceholderName
    //ExSummary:Shows how to use a building block's contents as a custom placeholder text for a structured document tag. 
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert a plain text structured document tag of the "PlainText" type, which will function as a text box.
    // The contents that it will display by default are a "Click here to enter text." prompt.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Inline);
    
    // We can get the tag to display the contents of a building block instead of the default text.
    // First, add a building block with contents to the glossary document.
    System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossaryDoc = doc->get_GlossaryDocument();
    
    auto substituteBlock = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    substituteBlock->set_Name(u"Custom Placeholder");
    substituteBlock->AppendChild<System::SharedPtr<Aspose::Words::Section>>(System::MakeObject<Aspose::Words::Section>(glossaryDoc));
    substituteBlock->get_FirstSection()->AppendChild<System::SharedPtr<Aspose::Words::Body>>(System::MakeObject<Aspose::Words::Body>(glossaryDoc));
    substituteBlock->get_FirstSection()->get_Body()->AppendParagraph(u"Custom placeholder text.");
    
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(substituteBlock);
    
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
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertNode(tag);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.PlaceholderBuildingBlock.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.PlaceholderBuildingBlock.docx");
    tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    substituteBlock = System::ExplicitCast<Aspose::Words::BuildingBlocks::BuildingBlock>(doc->get_GlossaryDocument()->GetChild(Aspose::Words::NodeType::BuildingBlock, 0, true));
    
    ASSERT_EQ(u"Custom Placeholder", substituteBlock->get_Name());
    ASPOSE_ASSERT_EQ(isShowingPlaceholderText, tag->get_IsShowingPlaceholderText());
    ASPOSE_ASSERT_EQ(substituteBlock, tag->get_Placeholder());
    ASSERT_EQ(substituteBlock->get_Name(), tag->get_PlaceholderName());
}

namespace gtest_test
{

using ExStructuredDocumentTag_PlaceholderBuildingBlock_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExStructuredDocumentTag::PlaceholderBuildingBlock)>::type;

struct ExStructuredDocumentTag_PlaceholderBuildingBlock : public ExStructuredDocumentTag, public Aspose::Words::ApiExamples::ExStructuredDocumentTag, public ::testing::WithParamInterface<ExStructuredDocumentTag_PlaceholderBuildingBlock_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExStructuredDocumentTag_PlaceholderBuildingBlock, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PlaceholderBuildingBlock(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExStructuredDocumentTag_PlaceholderBuildingBlock, ::testing::ValuesIn(ExStructuredDocumentTag_PlaceholderBuildingBlock::TestCases()));

} // namespace gtest_test

void ExStructuredDocumentTag::Lock()
{
    //ExStart
    //ExFor:StructuredDocumentTag.LockContentControl
    //ExFor:StructuredDocumentTag.LockContents
    //ExFor:IStructuredDocumentTag.LockContentControl
    //ExFor:IStructuredDocumentTag.LockContents
    //ExSummary:Shows how to apply editing restrictions to structured document tags.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a plain text structured document tag, which acts as a text box that prompts the user to fill it in.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Inline);
    
    // Set the "LockContents" property to "true" to prohibit the user from editing this text box's contents.
    tag->set_LockContents(true);
    builder->Write(u"The contents of this structured document tag cannot be edited: ");
    builder->InsertNode(tag);
    
    tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Inline);
    
    // Set the "LockContentControl" property to "true" to prohibit the user from
    // deleting this structured document tag manually in Microsoft Word.
    tag->set_LockContentControl(true);
    
    builder->InsertParagraph();
    builder->Write(u"This structured document tag cannot be deleted but its contents can be edited: ");
    builder->InsertNode(tag);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.Lock.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.Lock.docx");
    tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    
    ASSERT_TRUE(tag->get_LockContents());
    ASSERT_FALSE(tag->get_LockContentControl());
    
    tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 1, true));
    
    ASSERT_FALSE(tag->get_LockContents());
    ASSERT_TRUE(tag->get_LockContentControl());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, Lock)
{
    s_instance->Lock();
}

} // namespace gtest_test

void ExStructuredDocumentTag::ListItemCollection()
{
    //ExStart
    //ExFor:SdtListItem
    //ExFor:SdtListItem.#ctor(String)
    //ExFor:SdtListItem.#ctor(String,String)
    //ExFor:SdtListItem.DisplayText
    //ExFor:SdtListItem.Value
    //ExFor:SdtListItemCollection
    //ExFor:SdtListItemCollection.Add(SdtListItem)
    //ExFor:SdtListItemCollection.Clear
    //ExFor:SdtListItemCollection.Count
    //ExFor:SdtListItemCollection.GetEnumerator
    //ExFor:SdtListItemCollection.Item(Int32)
    //ExFor:SdtListItemCollection.RemoveAt(Int32)
    //ExFor:SdtListItemCollection.SelectedValue
    //ExFor:StructuredDocumentTag.ListItems
    //ExSummary:Shows how to work with drop down-list structured document tags.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::DropDownList, Aspose::Words::Markup::MarkupLevel::Block);
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(tag);
    
    // A drop-down list structured document tag is a form that allows the user to
    // select an option from a list by left-clicking and opening the form in Microsoft Word.
    // The "ListItems" property contains all list items, and each list item is an "SdtListItem".
    System::SharedPtr<Aspose::Words::Markup::SdtListItemCollection> listItems = tag->get_ListItems();
    listItems->Add(System::MakeObject<Aspose::Words::Markup::SdtListItem>(u"Value 1"));
    
    ASSERT_EQ(listItems->idx_get(0)->get_DisplayText(), listItems->idx_get(0)->get_Value());
    
    // Add 3 more list items. Initialize these items using a different constructor to the first item
    // to display strings that are different from their values.
    listItems->Add(System::MakeObject<Aspose::Words::Markup::SdtListItem>(u"Item 2", u"Value 2"));
    listItems->Add(System::MakeObject<Aspose::Words::Markup::SdtListItem>(u"Item 3", u"Value 3"));
    listItems->Add(System::MakeObject<Aspose::Words::Markup::SdtListItem>(u"Item 4", u"Value 4"));
    
    ASSERT_EQ(4, listItems->get_Count());
    
    // The drop-down list is displaying the first item. Assign a different list item to the "SelectedValue" to display it.
    listItems->set_SelectedValue(listItems->idx_get(3));
    
    ASSERT_EQ(u"Value 4", listItems->get_SelectedValue()->get_Value());
    
    // Enumerate over the collection and print each element.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Markup::SdtListItem>>> enumerator = listItems->GetEnumerator();
        while (enumerator->MoveNext())
        {
            if (enumerator->get_Current() != nullptr)
            {
                std::cout << System::String::Format(u"List item: {0}, value: {1}", enumerator->get_Current()->get_DisplayText(), enumerator->get_Current()->get_Value()) << std::endl;
            }
        }
    }
    
    // Remove the last list item. 
    listItems->RemoveAt(3);
    
    ASSERT_EQ(3, listItems->get_Count());
    
    // Since our drop-down control is set to display the removed item by default, give it an item to display which exists.
    listItems->set_SelectedValue(listItems->idx_get(1));
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.ListItemCollection.docx");
    
    // Use the "Clear" method to empty the entire drop-down item collection at once.
    listItems->Clear();
    
    ASSERT_EQ(0, listItems->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, ListItemCollection)
{
    s_instance->ListItemCollection();
}

} // namespace gtest_test

void ExStructuredDocumentTag::CreatingCustomXml()
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
    //ExFor:IStructuredDocumentTag.XmlMapping
    //ExFor:XmlMapping.SetMapping(CustomXmlPart, String, String)
    //ExSummary:Shows how to create a structured document tag with custom XML data.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Construct an XML part that contains data and add it to the document's collection.
    // If we enable the "Developer" tab in Microsoft Word,
    // we can find elements from this collection in the "XML Mapping Pane", along with a few default elements.
    System::String xmlPartId = System::Guid::NewGuid().ToString(u"B");
    System::String xmlPartContent = u"<root><text>Hello world!</text></root>";
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);
    
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
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPart> xmlPartClone = xmlPart->Clone();
    xmlPartClone->set_Id(System::Guid::NewGuid().ToString(u"B"));
    doc->get_CustomXmlParts()->Add(xmlPartClone);
    
    ASSERT_EQ(2, doc->get_CustomXmlParts()->get_Count());
    
    // Iterate through the collection and print the contents of each part.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Markup::CustomXmlPart>>> enumerator = doc->get_CustomXmlParts()->GetEnumerator();
        int32_t index = 0;
        while (enumerator->MoveNext())
        {
            std::cout << System::String::Format(u"XML part index {0}, ID: {1}", index, enumerator->get_Current()->get_Id()) << std::endl;
            std::cout << System::String::Format(u"\tContent: {0}", System::Text::Encoding::get_UTF8()->GetString(enumerator->get_Current()->get_Data())) << std::endl;
            index++;
        }
    }
    
    // Use the "RemoveAt" method to remove the cloned part by index.
    doc->get_CustomXmlParts()->RemoveAt(1);
    
    ASSERT_EQ(1, doc->get_CustomXmlParts()->get_Count());
    
    // Clone the XML parts collection, and then use the "Clear" method to remove all its elements at once.
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPartCollection> customXmlParts = doc->get_CustomXmlParts()->Clone();
    customXmlParts->Clear();
    
    // Create a structured document tag that will display our part's contents and insert it into the document body.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Block);
    tag->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[1]", System::String::Empty);
    
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(tag);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.CustomXml.docx");
    //ExEnd
    
    ASSERT_TRUE(Aspose::Words::ApiExamples::DocumentHelper::CompareDocs(get_ArtifactsDir() + u"StructuredDocumentTag.CustomXml.docx", get_GoldsDir() + u"StructuredDocumentTag.CustomXml Gold.docx"));
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.CustomXml.docx");
    xmlPart = doc->get_CustomXmlParts()->idx_get(0);
    
    ASSERT_NO_THROW(static_cast<std::function<void()>>([&xmlPart]() -> void
    {
        System::Guid::Parse(xmlPart->get_Id());
    })());
    ASSERT_EQ(u"<root><text>Hello world!</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));
    ASSERT_EQ(u"http://www.w3.org/2001/XMLSchema", xmlPart->get_Schemas()->idx_get(0));
    
    tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    ASSERT_EQ(u"Hello world!", tag->GetText().Trim());
    ASSERT_EQ(u"/root[1]/text[1]", tag->get_XmlMapping()->get_XPath());
    ASSERT_EQ(System::String::Empty, tag->get_XmlMapping()->get_PrefixMappings());
    ASSERT_EQ(xmlPart->get_DataChecksum(), tag->get_XmlMapping()->get_CustomXmlPart()->get_DataChecksum());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, CreatingCustomXml)
{
    s_instance->CreatingCustomXml();
}

} // namespace gtest_test

void ExStructuredDocumentTag::DataChecksum()
{
    //ExStart
    //ExFor:CustomXmlPart.DataChecksum
    //ExSummary:Shows how the checksum is calculated in a runtime.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto richText = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::RichText, Aspose::Words::Markup::MarkupLevel::Block);
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(richText);
    
    // The checksum is read-only and computed using the data of the corresponding custom XML data part.
    richText->get_XmlMapping()->SetMapping(doc->get_CustomXmlParts()->Add(System::ObjectExt::ToString(System::Guid::NewGuid()), u"<root><text>ContentControl</text></root>"), u"/root/text", u"");
    
    int64_t checksum = richText->get_XmlMapping()->get_CustomXmlPart()->get_DataChecksum();
    std::cout << checksum << std::endl;
    
    richText->get_XmlMapping()->SetMapping(doc->get_CustomXmlParts()->Add(System::ObjectExt::ToString(System::Guid::NewGuid()), u"<root><text>Updated ContentControl</text></root>"), u"/root/text", u"");
    
    int64_t updatedChecksum = richText->get_XmlMapping()->get_CustomXmlPart()->get_DataChecksum();
    std::cout << updatedChecksum << std::endl;
    
    // We changed the XmlPart of the tag, and the checksum was updated at runtime.
    ASSERT_NE(checksum, updatedChecksum);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, DataChecksum)
{
    s_instance->DataChecksum();
}

} // namespace gtest_test

void ExStructuredDocumentTag::XmlMapping()
{
    //ExStart
    //ExFor:XmlMapping
    //ExFor:XmlMapping.CustomXmlPart
    //ExFor:XmlMapping.Delete
    //ExFor:XmlMapping.IsMapped
    //ExFor:XmlMapping.PrefixMappings
    //ExFor:XmlMapping.XPath
    //ExSummary:Shows how to set XML mappings for custom XML parts.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
    System::String xmlPartId = System::Guid::NewGuid().ToString(u"B");
    System::String xmlPartContent = u"<root><text>Text element #1</text><text>Text element #2</text></root>";
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);
    
    ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));
    
    // Create a structured document tag that will display the contents of our CustomXmlPart.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Block);
    
    // Set a mapping for our structured document tag. This mapping will instruct
    // our structured document tag to display a portion of the XML part's text contents that the XPath points to.
    // In this case, it will be contents of the the second "<text>" element of the first "<root>" element: "Text element #2".
    tag->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[2]", u"xmlns:ns='http://www.w3.org/2001/XMLSchema'");
    
    ASSERT_TRUE(tag->get_XmlMapping()->get_IsMapped());
    ASPOSE_ASSERT_EQ(xmlPart, tag->get_XmlMapping()->get_CustomXmlPart());
    ASSERT_EQ(u"/root[1]/text[2]", tag->get_XmlMapping()->get_XPath());
    ASSERT_EQ(u"xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag->get_XmlMapping()->get_PrefixMappings());
    
    // Add the structured document tag to the document to display the content from our custom part.
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(tag);
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.XmlMapping.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.XmlMapping.docx");
    xmlPart = doc->get_CustomXmlParts()->idx_get(0);
    
    ASSERT_NO_THROW(static_cast<std::function<void()>>([&xmlPart]() -> void
    {
        System::Guid::Parse(xmlPart->get_Id());
    })());
    ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));
    
    tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    ASSERT_EQ(u"Text element #2", tag->GetText().Trim());
    ASSERT_EQ(u"/root[1]/text[2]", tag->get_XmlMapping()->get_XPath());
    ASSERT_EQ(u"xmlns:ns='http://www.w3.org/2001/XMLSchema'", tag->get_XmlMapping()->get_PrefixMappings());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, XmlMapping)
{
    s_instance->XmlMapping();
}

} // namespace gtest_test

void ExStructuredDocumentTag::StructuredDocumentTagRangeStartXmlMapping()
{
    //ExStart
    //ExFor:StructuredDocumentTagRangeStart.XmlMapping
    //ExSummary:Shows how to set XML mappings for the range start of a structured document tag.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Multi-section structured document tags.docx");
    
    // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
    System::String xmlPartId = System::Guid::NewGuid().ToString(u"B");
    System::String xmlPartContent = u"<root><text>Text element #1</text><text>Text element #2</text></root>";
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);
    
    ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));
    
    // Create a structured document tag that will display the contents of our CustomXmlPart in the document.
    auto sdtRangeStart = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, 0, true));
    
    // If we set a mapping for our structured document tag,
    // it will only display a portion of the CustomXmlPart that the XPath points to.
    // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
    sdtRangeStart->get_XmlMapping()->SetMapping(xmlPart, u"/root[1]/text[2]", nullptr);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
    xmlPart = doc->get_CustomXmlParts()->idx_get(0);
    
    ASSERT_NO_THROW(static_cast<std::function<void()>>([&xmlPart]() -> void
    {
        System::Guid::Parse(xmlPart->get_Id());
    })());
    ASSERT_EQ(u"<root><text>Text element #1</text><text>Text element #2</text></root>", System::Text::Encoding::get_UTF8()->GetString(xmlPart->get_Data()));
    
    sdtRangeStart = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, 0, true));
    ASSERT_EQ(u"/root[1]/text[2]", sdtRangeStart->get_XmlMapping()->get_XPath());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, StructuredDocumentTagRangeStartXmlMapping)
{
    s_instance->StructuredDocumentTagRangeStartXmlMapping();
}

} // namespace gtest_test

void ExStructuredDocumentTag::CustomXmlSchemaCollection()
{
    //ExStart
    //ExFor:CustomXmlSchemaCollection
    //ExFor:CustomXmlSchemaCollection.Add(String)
    //ExFor:CustomXmlSchemaCollection.Clear
    //ExFor:CustomXmlSchemaCollection.Clone
    //ExFor:CustomXmlSchemaCollection.Count
    //ExFor:CustomXmlSchemaCollection.GetEnumerator
    //ExFor:CustomXmlSchemaCollection.IndexOf(String)
    //ExFor:CustomXmlSchemaCollection.Item(Int32)
    //ExFor:CustomXmlSchemaCollection.Remove(String)
    //ExFor:CustomXmlSchemaCollection.RemoveAt(Int32)
    //ExSummary:Shows how to work with an XML schema collection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::String xmlPartId = System::Guid::NewGuid().ToString(u"B");
    System::String xmlPartContent = u"<root><text>Hello, World!</text></root>";
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(xmlPartId, xmlPartContent);
    
    // Add an XML schema association.
    xmlPart->get_Schemas()->Add(u"http://www.w3.org/2001/XMLSchema");
    
    // Clone the custom XML part's XML schema association collection,
    // and then add a couple of new schemas to the clone.
    System::SharedPtr<Aspose::Words::Markup::CustomXmlSchemaCollection> schemas = xmlPart->get_Schemas()->Clone();
    schemas->Add(u"http://www.w3.org/2001/XMLSchema-instance");
    schemas->Add(u"http://schemas.microsoft.com/office/2006/metadata/contentType");
    
    ASSERT_EQ(3, schemas->get_Count());
    ASSERT_EQ(2, schemas->IndexOf(u"http://schemas.microsoft.com/office/2006/metadata/contentType"));
    
    // Enumerate the schemas and print each element.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::String>> enumerator = schemas->GetEnumerator();
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

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, CustomXmlSchemaCollection)
{
    s_instance->CustomXmlSchemaCollection();
}

} // namespace gtest_test

void ExStructuredDocumentTag::CustomXmlPartStoreItemIdReadOnly()
{
    //ExStart
    //ExFor:XmlMapping.StoreItemId
    //ExSummary:Shows how to get the custom XML data identifier of an XML part.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Custom XML part in structured document tag.docx");
    
    // Structured document tags have IDs in the form of GUIDs.
    auto tag = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    
    ASSERT_EQ(u"{F3029283-4FF8-4DD2-9F31-395F19ACEE85}", tag->get_XmlMapping()->get_StoreItemId());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, CustomXmlPartStoreItemIdReadOnly)
{
    s_instance->CustomXmlPartStoreItemIdReadOnly();
}

} // namespace gtest_test

void ExStructuredDocumentTag::CustomXmlPartStoreItemIdReadOnlyNull()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto sdtCheckBox = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::Checkbox, Aspose::Words::Markup::MarkupLevel::Inline);
    sdtCheckBox->set_Checked(true);
    
    builder->InsertNode(sdtCheckBox);
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    auto sdt = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    std::cout << (System::String(u"The Id of your custom xml part is: ") + sdt->get_XmlMapping()->get_StoreItemId()) << std::endl;
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, CustomXmlPartStoreItemIdReadOnlyNull)
{
    s_instance->CustomXmlPartStoreItemIdReadOnlyNull();
}

} // namespace gtest_test

void ExStructuredDocumentTag::ClearTextFromStructuredDocumentTags()
{
    //ExStart
    //ExFor:StructuredDocumentTag.Clear
    //ExSummary:Shows how to delete contents of structured document tag elements.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a plain text structured document tag, and then append it to the document.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Block);
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(tag);
    
    // This structured document tag, which is in the form of a text box, already displays placeholder text.
    ASSERT_EQ(u"Click here to enter text.", tag->GetText().Trim());
    ASSERT_TRUE(tag->get_IsShowingPlaceholderText());
    
    // Create a building block with text contents.
    System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossaryDoc = doc->get_GlossaryDocument();
    auto substituteBlock = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    substituteBlock->set_Name(u"My placeholder");
    substituteBlock->AppendChild<System::SharedPtr<Aspose::Words::Section>>(System::MakeObject<Aspose::Words::Section>(glossaryDoc));
    substituteBlock->get_FirstSection()->EnsureMinimum();
    substituteBlock->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(glossaryDoc, u"Custom placeholder text."));
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(substituteBlock);
    
    // Set the structured document tag's "PlaceholderName" property to our building block's name to get
    // the structured document tag to display the contents of the building block in place of the original default text.
    tag->set_PlaceholderName(u"My placeholder");
    
    ASSERT_EQ(u"Custom placeholder text.", tag->GetText().Trim());
    ASSERT_TRUE(tag->get_IsShowingPlaceholderText());
    
    // Edit the text of the structured document tag and hide the placeholder text.
    auto run = System::ExplicitCast<Aspose::Words::Run>(tag->GetChild(Aspose::Words::NodeType::Run, 0, true));
    run->set_Text(u"New text.");
    tag->set_IsShowingPlaceholderText(false);
    
    ASSERT_EQ(u"New text.", tag->GetText().Trim());
    
    // Use the "Clear" method to clear this structured document tag's contents and display the placeholder again.
    tag->Clear();
    
    ASSERT_TRUE(tag->get_IsShowingPlaceholderText());
    ASSERT_EQ(u"Custom placeholder text.", tag->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, ClearTextFromStructuredDocumentTags)
{
    s_instance->ClearTextFromStructuredDocumentTags();
}

} // namespace gtest_test

void ExStructuredDocumentTag::AccessToBuildingBlockPropertiesFromDocPartObjSdt()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags with building blocks.docx");
    
    auto docPartObjSdt = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    
    ASSERT_EQ(Aspose::Words::Markup::SdtType::DocPartObj, docPartObjSdt->get_SdtType());
    ASSERT_EQ(u"Table of Contents", docPartObjSdt->get_BuildingBlockGallery());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, AccessToBuildingBlockPropertiesFromDocPartObjSdt)
{
    s_instance->AccessToBuildingBlockPropertiesFromDocPartObjSdt();
}

} // namespace gtest_test

void ExStructuredDocumentTag::AccessToBuildingBlockPropertiesFromPlainTextSdt()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags with building blocks.docx");
    
    auto plainTextSdt = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 1, true));
    
    ASSERT_EQ(Aspose::Words::Markup::SdtType::PlainText, plainTextSdt->get_SdtType());
    ASSERT_THROW(static_cast<std::function<void()>>([&plainTextSdt]() -> void
    {
        System::String _ = plainTextSdt->get_BuildingBlockGallery();
    })(), System::InvalidOperationException) << "BuildingBlockType is only accessible for BuildingBlockGallery SDT type.";
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, AccessToBuildingBlockPropertiesFromPlainTextSdt)
{
    s_instance->AccessToBuildingBlockPropertiesFromPlainTextSdt();
}

} // namespace gtest_test

void ExStructuredDocumentTag::BuildingBlockCategories()
{
    //ExStart
    //ExFor:StructuredDocumentTag.BuildingBlockCategory
    //ExFor:StructuredDocumentTag.BuildingBlockGallery
    //ExSummary:Shows how to insert a structured document tag as a building block, and set its category and gallery.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto buildingBlockSdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::BuildingBlockGallery, Aspose::Words::Markup::MarkupLevel::Block);
    buildingBlockSdt->set_BuildingBlockCategory(u"Built-in");
    buildingBlockSdt->set_BuildingBlockGallery(u"Table of Contents");
    
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(buildingBlockSdt);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.BuildingBlockCategories.docx");
    //ExEnd
    
    buildingBlockSdt = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTag>(doc->get_FirstSection()->get_Body()->GetChild(Aspose::Words::NodeType::StructuredDocumentTag, 0, true));
    
    ASSERT_EQ(Aspose::Words::Markup::SdtType::BuildingBlockGallery, buildingBlockSdt->get_SdtType());
    ASSERT_EQ(u"Table of Contents", buildingBlockSdt->get_BuildingBlockGallery());
    ASSERT_EQ(u"Built-in", buildingBlockSdt->get_BuildingBlockCategory());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, BuildingBlockCategories)
{
    s_instance->BuildingBlockCategories();
}

} // namespace gtest_test

void ExStructuredDocumentTag::UpdateSdtContent()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert a drop-down list structured document tag.
    auto tag = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::DropDownList, Aspose::Words::Markup::MarkupLevel::Block);
    tag->get_ListItems()->Add(System::MakeObject<Aspose::Words::Markup::SdtListItem>(u"Value 1"));
    tag->get_ListItems()->Add(System::MakeObject<Aspose::Words::Markup::SdtListItem>(u"Value 2"));
    tag->get_ListItems()->Add(System::MakeObject<Aspose::Words::Markup::SdtListItem>(u"Value 3"));
    
    // The drop-down list currently displays "Choose an item" as the default text.
    // Set the "SelectedValue" property to one of the list items to get the tag to
    // display that list item's value instead of the default text.
    tag->get_ListItems()->set_SelectedValue(tag->get_ListItems()->idx_get(1));
    
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(tag);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.UpdateSdtContent.pdf");
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, UpdateSdtContent)
{
    s_instance->UpdateSdtContent();
}

} // namespace gtest_test

void ExStructuredDocumentTag::FillTableUsingRepeatingSectionItem()
{
    //ExStart
    //ExFor:SdtType
    //ExSummary:Shows how to fill a table with data from in an XML part.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(u"Books", System::String(u"<books>") + u"<book>" + u"<title>Everyday Italian</title>" + u"<author>Giada De Laurentiis</author>" + u"</book>" + u"<book>" + u"<title>The C Programming Language</title>" + u"<author>Brian W. Kernighan, Dennis M. Ritchie</author>" + u"</book>" + u"<book>" + u"<title>Learning XML</title>" + u"<author>Erik T. Ray</author>" + u"</book>" + u"</books>");
    
    // Create headers for data from the XML content.
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Title");
    builder->InsertCell();
    builder->Write(u"Author");
    builder->EndRow();
    builder->EndTable();
    
    // Create a table with a repeating section inside.
    auto repeatingSectionSdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::RepeatingSection, Aspose::Words::Markup::MarkupLevel::Row);
    repeatingSectionSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book", System::String::Empty);
    table->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(repeatingSectionSdt);
    
    // Add repeating section item inside the repeating section and mark it as a row.
    // This table will have a row for each element that we can find in the XML document
    // using the "/books[1]/book" XPath, of which there are three.
    auto repeatingSectionItemSdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::RepeatingSectionItem, Aspose::Words::Markup::MarkupLevel::Row);
    repeatingSectionSdt->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(repeatingSectionItemSdt);
    
    auto row = System::MakeObject<Aspose::Words::Tables::Row>(doc);
    repeatingSectionItemSdt->AppendChild<System::SharedPtr<Aspose::Words::Tables::Row>>(row);
    
    // Map XML data with created table cells for the title and author of each book.
    auto titleSdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Cell);
    titleSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/title[1]", System::String::Empty);
    row->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(titleSdt);
    
    auto authorSdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Cell);
    authorSdt->get_XmlMapping()->SetMapping(xmlPart, u"/books[1]/book[1]/author[1]", System::String::Empty);
    row->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(authorSdt);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.RepeatingSectionItem.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.RepeatingSectionItem.docx");
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>> tags = doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTag, true)->LINQ_OfType<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> >()->LINQ_ToList();
    
    ASSERT_EQ(u"/books[1]/book", tags->idx_get(0)->get_XmlMapping()->get_XPath());
    ASSERT_EQ(System::String::Empty, tags->idx_get(0)->get_XmlMapping()->get_PrefixMappings());
    
    ASSERT_EQ(System::String::Empty, tags->idx_get(1)->get_XmlMapping()->get_XPath());
    ASSERT_EQ(System::String::Empty, tags->idx_get(1)->get_XmlMapping()->get_PrefixMappings());
    
    ASSERT_EQ(u"/books[1]/book[1]/title[1]", tags->idx_get(2)->get_XmlMapping()->get_XPath());
    ASSERT_EQ(System::String::Empty, tags->idx_get(2)->get_XmlMapping()->get_PrefixMappings());
    
    ASSERT_EQ(u"/books[1]/book[1]/author[1]", tags->idx_get(3)->get_XmlMapping()->get_XPath());
    ASSERT_EQ(System::String::Empty, tags->idx_get(3)->get_XmlMapping()->get_PrefixMappings());
    
    ASSERT_EQ(System::String(u"Title\u0007Author\u0007\u0007") + u"Everyday Italian\u0007Giada De Laurentiis\u0007\u0007" + u"The C Programming Language\u0007Brian W. Kernighan, Dennis M. Ritchie\u0007\u0007" + u"Learning XML\u0007Erik T. Ray\u0007\u0007", doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0)->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, FillTableUsingRepeatingSectionItem)
{
    s_instance->FillTableUsingRepeatingSectionItem();
}

} // namespace gtest_test

void ExStructuredDocumentTag::CustomXmlPart()
{
    System::String xmlString = System::String(u"<?xml version=\"1.0\"?>") + u"<Company>" + u"<Employee id=\"1\">" + u"<FirstName>John</FirstName>" + u"<LastName>Doe</LastName>" + u"</Employee>" + u"<Employee id=\"2\">" + u"<FirstName>Jane</FirstName>" + u"<LastName>Doe</LastName>" + u"</Employee>" + u"</Company>";
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Insert the full XML document as a custom document part.
    // We can find the mapping for this part in Microsoft Word via "Developer" -> "XML Mapping Pane", if it is enabled.
    System::SharedPtr<Aspose::Words::Markup::CustomXmlPart> xmlPart = doc->get_CustomXmlParts()->Add(System::Guid::NewGuid().ToString(u"B"), xmlString);
    
    // Create a structured document tag, which will use an XPath to refer to a single element from the XML.
    auto sdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::PlainText, Aspose::Words::Markup::MarkupLevel::Block);
    sdt->get_XmlMapping()->SetMapping(xmlPart, u"Company//Employee[@id='2']/FirstName", u"");
    
    // Add the StructuredDocumentTag to the document to display the element in the text.
    doc->get_FirstSection()->get_Body()->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(sdt);
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, CustomXmlPart)
{
    s_instance->CustomXmlPart();
}

} // namespace gtest_test

void ExStructuredDocumentTag::MultiSectionTags()
{
    //ExStart
    //ExFor:StructuredDocumentTagRangeStart
    //ExFor:IStructuredDocumentTag.Id
    //ExFor:StructuredDocumentTagRangeStart.Id
    //ExFor:StructuredDocumentTagRangeStart.Title
    //ExFor:StructuredDocumentTagRangeStart.PlaceholderName
    //ExFor:StructuredDocumentTagRangeStart.IsShowingPlaceholderText
    //ExFor:StructuredDocumentTagRangeStart.LockContentControl
    //ExFor:StructuredDocumentTagRangeStart.LockContents
    //ExFor:IStructuredDocumentTag.Level
    //ExFor:StructuredDocumentTagRangeStart.Level
    //ExFor:StructuredDocumentTagRangeStart.RangeEnd
    //ExFor:IStructuredDocumentTag.Color
    //ExFor:StructuredDocumentTagRangeStart.Color
    //ExFor:StructuredDocumentTagRangeStart.SdtType
    //ExFor:StructuredDocumentTagRangeStart.WordOpenXML
    //ExFor:StructuredDocumentTagRangeStart.Tag
    //ExFor:StructuredDocumentTagRangeEnd
    //ExFor:StructuredDocumentTagRangeEnd.Id
    //ExSummary:Shows how to get the properties of multi-section structured document tags.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Multi-section structured document tags.docx");
    
    auto rangeStartTag = System::AsCast<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, true)->idx_get(0));
    auto rangeEndTag = System::AsCast<Aspose::Words::Markup::StructuredDocumentTagRangeEnd>(doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTagRangeEnd, true)->idx_get(0));
    
    ASSERT_EQ(rangeStartTag->get_Id(), rangeEndTag->get_Id());
    //ExSkip
    ASSERT_EQ(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, rangeStartTag->get_NodeType());
    //ExSkip
    ASSERT_EQ(Aspose::Words::NodeType::StructuredDocumentTagRangeEnd, rangeEndTag->get_NodeType());
    //ExSkip
    
    std::cout << "StructuredDocumentTagRangeStart values:" << std::endl;
    std::cout << System::String::Format(u"\t|Id: {0}", rangeStartTag->get_Id()) << std::endl;
    std::cout << System::String::Format(u"\t|Title: {0}", rangeStartTag->get_Title()) << std::endl;
    std::cout << System::String::Format(u"\t|PlaceholderName: {0}", rangeStartTag->get_PlaceholderName()) << std::endl;
    std::cout << System::String::Format(u"\t|IsShowingPlaceholderText: {0}", rangeStartTag->get_IsShowingPlaceholderText()) << std::endl;
    std::cout << System::String::Format(u"\t|LockContentControl: {0}", rangeStartTag->get_LockContentControl()) << std::endl;
    std::cout << System::String::Format(u"\t|LockContents: {0}", rangeStartTag->get_LockContents()) << std::endl;
    std::cout << System::String::Format(u"\t|Level: {0}", rangeStartTag->get_Level()) << std::endl;
    std::cout << System::String::Format(u"\t|NodeType: {0}", rangeStartTag->get_NodeType()) << std::endl;
    std::cout << System::String::Format(u"\t|RangeEnd: {0}", rangeStartTag->get_RangeEnd()) << std::endl;
    std::cout << System::String::Format(u"\t|Color: {0}", rangeStartTag->get_Color().ToArgb()) << std::endl;
    std::cout << System::String::Format(u"\t|SdtType: {0}", rangeStartTag->get_SdtType()) << std::endl;
    std::cout << System::String::Format(u"\t|FlatOpcContent: {0}", rangeStartTag->get_WordOpenXML()) << std::endl;
    std::cout << System::String::Format(u"\t|Tag: {0}\n", rangeStartTag->get_Tag()) << std::endl;
    
    std::cout << "StructuredDocumentTagRangeEnd values:" << std::endl;
    std::cout << System::String::Format(u"\t|Id: {0}", rangeEndTag->get_Id()) << std::endl;
    std::cout << System::String::Format(u"\t|NodeType: {0}", rangeEndTag->get_NodeType()) << std::endl;
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, MultiSectionTags)
{
    s_instance->MultiSectionTags();
}

} // namespace gtest_test

void ExStructuredDocumentTag::SdtChildNodes()
{
    //ExStart
    //ExFor:StructuredDocumentTagRangeStart.GetChildNodes(NodeType, bool)
    //ExSummary:Shows how to get child nodes of StructuredDocumentTagRangeStart.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Multi-section structured document tags.docx");
    auto tag = System::AsCast<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc->GetChildNodes(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, true)->idx_get(0));
    
    std::cout << "StructuredDocumentTagRangeStart values:" << std::endl;
    std::cout << System::String::Format(u"\t|Child nodes count: {0}\n", tag->GetChildNodes(Aspose::Words::NodeType::Any, false)->get_Count()) << std::endl;
    
    for (auto&& node : System::IterateOver(tag->GetChildNodes(Aspose::Words::NodeType::Any, false)))
    {
        std::cout << System::String::Format(u"\t|Child node type: {0}", node->get_NodeType()) << std::endl;
    }
    
    for (auto&& node : System::IterateOver(tag->GetChildNodes(Aspose::Words::NodeType::Run, true)))
    {
        std::cout << System::String::Format(u"\t|Child node text: {0}", node->GetText()) << std::endl;
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, SdtChildNodes)
{
    s_instance->SdtChildNodes();
}

} // namespace gtest_test

void ExStructuredDocumentTag::SdtRangeExtendedMethods()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"StructuredDocumentTag element");
    
    System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagRangeStart> rangeStart = InsertStructuredDocumentTagRanges(doc);
    
    // Removes ranged structured document tag, but keeps content inside.
    rangeStart->RemoveSelfOnly();
    
    rangeStart = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, 0, false));
    ASPOSE_ASSERT_EQ(nullptr, rangeStart);
    
    auto rangeEnd = System::ExplicitCast<Aspose::Words::Markup::StructuredDocumentTagRangeEnd>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTagRangeEnd, 0, false));
    
    ASPOSE_ASSERT_EQ(nullptr, rangeEnd);
    ASSERT_EQ(u"StructuredDocumentTag element", doc->GetText().Trim());
    
    rangeStart = InsertStructuredDocumentTagRanges(doc);
    
    System::SharedPtr<Aspose::Words::Node> paragraphNode = rangeStart->LINQ_LastOrDefault();
    System::String actual = System::Default<System::String>();
    System::SharedPtr<Aspose::Words::Node> condExpression = paragraphNode;
    if (condExpression != nullptr)
    {
        actual = condExpression->GetText().Trim();
    }
    ASSERT_EQ(u"StructuredDocumentTag element", actual);
    
    // Removes ranged structured document tag and content inside.
    rangeStart->RemoveAllChildren();
    
    paragraphNode = rangeStart->LINQ_LastOrDefault();
    System::String actual2 = System::Default<System::String>();
    System::SharedPtr<Aspose::Words::Node> condExpression2 = paragraphNode;
    if (condExpression2 != nullptr)
    {
        actual2 = condExpression2->GetText();
    }
    ASPOSE_ASSERT_EQ(nullptr, actual2);
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, SdtRangeExtendedMethods)
{
    s_instance->SdtRangeExtendedMethods();
}

} // namespace gtest_test

void ExStructuredDocumentTag::GetSdt()
{
    //ExStart
    //ExFor:Range.StructuredDocumentTags
    //ExFor:StructuredDocumentTagCollection.Remove(int)
    //ExFor:StructuredDocumentTagCollection.RemoveAt(int)
    //ExSummary:Shows how to remove structured document tag.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags.docx");
    
    System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTagCollection> structuredDocumentTags = doc->get_Range()->get_StructuredDocumentTags();
    System::SharedPtr<Aspose::Words::Markup::IStructuredDocumentTag> sdt;
    for (int32_t i = 0; i < structuredDocumentTags->get_Count(); i++)
    {
        sdt = structuredDocumentTags->idx_get(i);
        std::cout << sdt->get_Title() << std::endl;
    }
    
    sdt = structuredDocumentTags->GetById(1691867797);
    ASSERT_EQ(1691867797, sdt->get_Id());
    
    ASSERT_EQ(5, structuredDocumentTags->get_Count());
    // Remove the structured document tag by Id.
    structuredDocumentTags->Remove(1691867797);
    // Remove the structured document tag at position 0.
    structuredDocumentTags->RemoveAt(0);
    ASSERT_EQ(3, structuredDocumentTags->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, GetSdt)
{
    s_instance->GetSdt();
}

} // namespace gtest_test

void ExStructuredDocumentTag::RangeSdt()
{
    //ExStart
    //ExFor:StructuredDocumentTagCollection
    //ExFor:StructuredDocumentTagCollection.GetById(int)
    //ExFor:StructuredDocumentTagCollection.GetByTitle(String)
    //ExFor:IStructuredDocumentTag.IsMultiSection
    //ExFor:IStructuredDocumentTag.Title
    //ExSummary:Shows how to get structured document tag.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags by id.docx");
    
    // Get the structured document tag by Id.
    System::SharedPtr<Aspose::Words::Markup::IStructuredDocumentTag> sdt = doc->get_Range()->get_StructuredDocumentTags()->GetById(1160505028);
    std::cout << System::Convert::ToString(sdt->get_IsMultiSection()) << std::endl;
    std::cout << sdt->get_Title() << std::endl;
    
    // Get the structured document tag or ranged tag by Title.
    sdt = doc->get_Range()->get_StructuredDocumentTags()->GetByTitle(u"Alias4");
    std::cout << sdt->get_Id() << std::endl;
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, RangeSdt)
{
    s_instance->RangeSdt();
}

} // namespace gtest_test

void ExStructuredDocumentTag::SdtAtRowLevel()
{
    //ExStart
    //ExFor:SdtType
    //ExSummary:Shows how to create group structured document tag at the Row level.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    
    // Create a Group structured document tag at the Row level.
    auto groupSdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::Group, Aspose::Words::Markup::MarkupLevel::Row);
    table->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(groupSdt);
    groupSdt->set_IsShowingPlaceholderText(false);
    groupSdt->RemoveAllChildren();
    
    // Create a child row of the structured document tag.
    auto row = System::MakeObject<Aspose::Words::Tables::Row>(doc);
    groupSdt->AppendChild<System::SharedPtr<Aspose::Words::Tables::Row>>(row);
    
    auto cell = System::MakeObject<Aspose::Words::Tables::Cell>(doc);
    row->AppendChild<System::SharedPtr<Aspose::Words::Tables::Cell>>(cell);
    
    builder->EndTable();
    
    // Insert cell contents.
    cell->EnsureMinimum();
    builder->MoveTo(cell->get_LastParagraph());
    builder->Write(u"Lorem ipsum dolor.");
    
    // Insert text after the table.
    builder->MoveTo(table->get_NextSibling());
    builder->Write(u"Nulla blandit nisi.");
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.SdtAtRowLevel.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, SdtAtRowLevel)
{
    s_instance->SdtAtRowLevel();
}

} // namespace gtest_test

void ExStructuredDocumentTag::IgnoreStructuredDocumentTags()
{
    //ExStart
    //ExFor:FindReplaceOptions.IgnoreStructuredDocumentTags
    //ExSummary:Shows how to ignore content of tags from replacement.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags.docx");
    
    // This paragraph contains SDT.
    auto p = System::ExplicitCast<Aspose::Words::Paragraph>(doc->get_FirstSection()->get_Body()->GetChild(Aspose::Words::NodeType::Paragraph, 2, true));
    System::String textToSearch = p->ToString(Aspose::Words::SaveFormat::Text).Trim();
    
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_IgnoreStructuredDocumentTags(true);
    doc->get_Range()->Replace(textToSearch, u"replacement", options);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.IgnoreStructuredDocumentTags.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"StructuredDocumentTag.IgnoreStructuredDocumentTags.docx");
    ASSERT_EQ(u"This document contains Structured Document Tags with text inside them\r\rRepeatingSection\rRichText\rreplacement", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, IgnoreStructuredDocumentTags)
{
    s_instance->IgnoreStructuredDocumentTags();
}

} // namespace gtest_test

void ExStructuredDocumentTag::Citation()
{
    //ExStart
    //ExFor:SdtType
    //ExSummary:Shows how to create a structured document tag of the Citation type.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto sdt = System::MakeObject<Aspose::Words::Markup::StructuredDocumentTag>(doc, Aspose::Words::Markup::SdtType::Citation, Aspose::Words::Markup::MarkupLevel::Inline);
    System::SharedPtr<Aspose::Words::Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    paragraph->AppendChild<System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag>>(sdt);
    
    // Create a Citation field.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveToParagraph(0, -1);
    builder->InsertField(u"CITATION Ath22 \\l 1033 ", u"(John Lennon, 2022)");
    
    // Move the field to the structured document tag.
    while (sdt->get_NextSibling() != nullptr)
    {
        sdt->AppendChild<System::SharedPtr<Aspose::Words::Node>>(sdt->get_NextSibling());
    }
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.Citation.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, Citation)
{
    s_instance->Citation();
}

} // namespace gtest_test

void ExStructuredDocumentTag::RangeStartWordOpenXmlMinimal()
{
    //ExStart:RangeStartWordOpenXmlMinimal
    //GistId:470c0da51e4317baae82ad9495747fed
    //ExFor:StructuredDocumentTagRangeStart.WordOpenXMLMinimal
    //ExSummary:Shows how to get minimal XML contained within the node in the FlatOpc format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Multi-section structured document tags.docx");
    auto tag = System::AsCast<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, 0, true));
    
    ASSERT_TRUE(tag->get_WordOpenXMLMinimal().Contains(u"<pkg:part pkg:name=\"/docProps/app.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\">"));
    ASSERT_FALSE(tag->get_WordOpenXMLMinimal().Contains(u"xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\""));
    //ExEnd:RangeStartWordOpenXmlMinimal
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, RangeStartWordOpenXmlMinimal)
{
    s_instance->RangeStartWordOpenXmlMinimal();
}

} // namespace gtest_test

void ExStructuredDocumentTag::RemoveSelfOnly()
{
    //ExStart:RemoveSelfOnly
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:IStructuredDocumentTag
    //ExFor:IStructuredDocumentTag.GetChildNodes(NodeType, bool)
    //ExFor:IStructuredDocumentTag.RemoveSelfOnly
    //ExSummary:Shows how to remove structured document tag, but keeps content inside.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Structured document tags.docx");
    
    // This collection provides a unified interface for accessing ranged and non-ranged structured tags. 
    System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::Markup::IStructuredDocumentTag>>> sdts = doc->get_Range()->get_StructuredDocumentTags()->LINQ_ToList();
    ASSERT_EQ(5, sdts->LINQ_Count());
    
    // Here we can get child nodes from the common interface of ranged and non-ranged structured tags.
    for (auto&& sdt : System::IterateOver(sdts))
    {
        if (sdt->GetChildNodes(Aspose::Words::NodeType::Any, false)->get_Count() > 0)
        {
            sdt->RemoveSelfOnly();
        }
    }
    
    sdts = doc->get_Range()->get_StructuredDocumentTags()->LINQ_ToList();
    ASSERT_EQ(0, sdts->LINQ_Count());
    //ExEnd:RemoveSelfOnly
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, RemoveSelfOnly)
{
    s_instance->RemoveSelfOnly();
}

} // namespace gtest_test

void ExStructuredDocumentTag::Appearance()
{
    //ExStart:Appearance
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:SdtAppearance
    //ExFor:StructuredDocumentTagRangeStart.Appearance
    //ExFor:IStructuredDocumentTag.Appearance
    //ExSummary:Shows how to show tag around content.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Multi-section structured document tags.docx");
    auto tag = System::AsCast<Aspose::Words::Markup::StructuredDocumentTagRangeStart>(doc->GetChild(Aspose::Words::NodeType::StructuredDocumentTagRangeStart, 0, true));
    
    if (tag->get_Appearance() == Aspose::Words::Markup::SdtAppearance::Hidden)
    {
        tag->set_Appearance(Aspose::Words::Markup::SdtAppearance::Tags);
    }
    //ExEnd:Appearance
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, Appearance)
{
    s_instance->Appearance();
}

} // namespace gtest_test

void ExStructuredDocumentTag::InsertStructuredDocumentTag()
{
    //ExStart:InsertStructuredDocumentTag
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBuilder.InsertStructuredDocumentTag(SdtType)
    //ExSummary:Shows how to simply insert structured document tag.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->MoveTo(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(3));
    // Note, that only following StructuredDocumentTag types are allowed for insertion:
    // SdtType.PlainText, SdtType.RichText, SdtType.Checkbox, SdtType.DropDownList,
    // SdtType.ComboBox, SdtType.Picture, SdtType.Date.
    // Markup level of inserted StructuredDocumentTag will be detected automatically and depends on position being inserted at.
    // Added StructuredDocumentTag will inherit paragraph and font formatting from cursor position.
    System::SharedPtr<Aspose::Words::Markup::StructuredDocumentTag> sdtPlain = builder->InsertStructuredDocumentTag(Aspose::Words::Markup::SdtType::PlainText);
    
    doc->Save(get_ArtifactsDir() + u"StructuredDocumentTag.InsertStructuredDocumentTag.docx");
    //ExEnd:InsertStructuredDocumentTag
}

namespace gtest_test
{

TEST_F(ExStructuredDocumentTag, InsertStructuredDocumentTag)
{
    s_instance->InsertStructuredDocumentTag();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
