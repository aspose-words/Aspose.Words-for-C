// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExLists.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTrailingCharacter.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelAlignment.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Lists;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1077218853u, ::Aspose::Words::ApiExamples::ExLists, ThisTypeBaseTypesInfo);

void ExLists::AddOutlineHeadingParagraphs(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list, System::String title)
{
    builder->get_ParagraphFormat()->ClearFormatting();
    builder->Writeln(title);
    
    for (int32_t i = 0; i < 9; i++)
    {
        builder->get_ListFormat()->set_List(list);
        builder->get_ListFormat()->set_ListLevelNumber(i);
        
        System::String styleName = System::String(u"Heading ") + (i + 1);
        builder->get_ParagraphFormat()->set_StyleName(styleName);
        builder->Writeln(styleName);
    }
    
    builder->get_ListFormat()->RemoveNumbers();
}

void ExLists::TestOutlineHeadingTemplates(System::SharedPtr<Aspose::Words::Document> doc)
{
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->idx_get(0);
    // Article section list template.
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"Article \u0000.", 0.0, Aspose::Words::NumberStyle::UppercaseRoman, list->get_ListLevels()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"Section \u0000.\u0001", 0.0, Aspose::Words::NumberStyle::LeadingZero, list->get_ListLevels()->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"(\u0002)", 14.4, Aspose::Words::NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(2));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"(\u0003)", 36.0, Aspose::Words::NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(3));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0004)", 28.8, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(4));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0005)", 36.0, Aspose::Words::NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(5));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0006)", 50.4, Aspose::Words::NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(6));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\a.", 50.4, Aspose::Words::NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(7));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\b.", 72.0, Aspose::Words::NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(8));
    
    list = doc->get_Lists()->idx_get(1);
    // Legal list template.
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(2));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(3));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(4));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(5));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(6));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(7));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a.\b", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(8));
    
    list = doc->get_Lists()->idx_get(2);
    // Numbered list template.
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.", 0.0, Aspose::Words::NumberStyle::UppercaseRoman, list->get_ListLevels()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0001.", 36.0, Aspose::Words::NumberStyle::UppercaseLetter, list->get_ListLevels()->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0002.", 72.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(2));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0003)", 108.0, Aspose::Words::NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(3));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"(\u0004)", 144.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(4));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"(\u0005)", 180.0, Aspose::Words::NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(5));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"(\u0006)", 216.0, Aspose::Words::NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(6));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"(\a)", 252.0, Aspose::Words::NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(7));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"(\b)", 288.0, Aspose::Words::NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(8));
    
    list = doc->get_Lists()->idx_get(3);
    // Chapter list template.
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"Chapter \u0000", 0.0, Aspose::Words::NumberStyle::Arabic, list->get_ListLevels()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(2));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(3));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(4));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(5));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(6));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(7));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"", 0.0, Aspose::Words::NumberStyle::None, list->get_ListLevels()->idx_get(8));
}

void ExLists::AddListSample(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list)
{
    builder->Writeln(System::String(u"Sample formatting of list with ListId:") + list->get_ListId());
    builder->get_ListFormat()->set_List(list);
    for (int32_t i = 0; i < list->get_ListLevels()->get_Count(); i++)
    {
        builder->get_ListFormat()->set_ListLevelNumber(i);
        builder->Writeln(System::String(u"Level ") + i);
    }
    
    builder->get_ListFormat()->RemoveNumbers();
    builder->Writeln();
}

void ExLists::TestPrintOutAllLists(System::SharedPtr<Aspose::Words::Document> listSourceDoc, System::SharedPtr<Aspose::Words::Document> outDoc)
{
    for (auto&& list : outDoc->get_Lists())
    {
        for (int32_t i = 0; i < list->get_ListLevels()->get_Count(); i++)
        {
            System::SharedPtr<Aspose::Words::Lists::ListLevel> expectedListLevel = listSourceDoc->get_Lists()->LINQ_First(static_cast<System::Func<System::SharedPtr<Aspose::Words::Lists::List>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Lists::List> l)>>([&list](System::SharedPtr<Aspose::Words::Lists::List> l) -> bool
            {
                return l->get_ListId() == list->get_ListId();
            })))->get_ListLevels()->idx_get(i);
            ASSERT_EQ(expectedListLevel->get_NumberFormat(), list->get_ListLevels()->idx_get(i)->get_NumberFormat());
            ASPOSE_ASSERT_EQ(expectedListLevel->get_NumberPosition(), list->get_ListLevels()->idx_get(i)->get_NumberPosition());
            ASSERT_EQ(expectedListLevel->get_NumberStyle(), list->get_ListLevels()->idx_get(i)->get_NumberStyle());
        }
    }
}


namespace gtest_test
{

class ExLists : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExLists> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExLists>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExLists> ExLists::s_instance;

} // namespace gtest_test

void ExLists::ApplyDefaultBulletsAndNumbers()
{
    //ExStart
    //ExFor:DocumentBuilder.ListFormat
    //ExFor:ListFormat.ApplyNumberDefault
    //ExFor:ListFormat.ApplyBulletDefault
    //ExFor:ListFormat.ListIndent
    //ExFor:ListFormat.ListOutdent
    //ExFor:ListFormat.RemoveNumbers
    //ExFor:ListFormat.ListLevelNumber
    //ExSummary:Shows how to create bulleted and numbered lists.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Aspose.Words main advantages are:");
    
    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Below are two types of lists that we can create with a document builder.
    // 1 -  A bulleted list:
    // This list will apply an indent and a bullet symbol ("•") before each paragraph.
    builder->get_ListFormat()->ApplyBulletDefault();
    builder->Writeln(u"Great performance");
    builder->Writeln(u"High reliability");
    builder->Writeln(u"Quality code and working");
    builder->Writeln(u"Wide variety of features");
    builder->Writeln(u"Easy to understand API");
    
    // End the bulleted list.
    builder->get_ListFormat()->RemoveNumbers();
    
    builder->InsertBreak(Aspose::Words::BreakType::ParagraphBreak);
    builder->Writeln(u"Aspose.Words allows:");
    
    // 2 -  A numbered list:
    // Numbered lists create a logical order for their paragraphs by numbering each item.
    builder->get_ListFormat()->ApplyNumberDefault();
    
    // This paragraph is the first item. The first item of a numbered list will have a "1." as its list item symbol.
    builder->Writeln(u"Opening documents from different formats:");
    
    ASSERT_EQ(0, builder->get_ListFormat()->get_ListLevelNumber());
    
    // Call the "ListIndent" method to increase the current list level,
    // which will start a new self-contained list, with a deeper indent, at the current item of the first list level.
    builder->get_ListFormat()->ListIndent();
    
    ASSERT_EQ(1, builder->get_ListFormat()->get_ListLevelNumber());
    
    // These are the first three list items of the second list level, which will maintain a count
    // independent of the count of the first list level. According to the current list format,
    // they will have symbols of "a.", "b.", and "c.".
    builder->Writeln(u"DOC");
    builder->Writeln(u"PDF");
    builder->Writeln(u"HTML");
    
    // Call the "ListOutdent" method to return to the previous list level.
    builder->get_ListFormat()->ListOutdent();
    
    ASSERT_EQ(0, builder->get_ListFormat()->get_ListLevelNumber());
    
    // These two paragraphs will continue the count of the first list level.
    // These items will have symbols of "2.", and "3."
    builder->Writeln(u"Processing documents");
    builder->Writeln(u"Saving documents in different formats:");
    
    // If we increase the list level to a level that we have added items to previously,
    // the nested list will be separate from the previous, and its numbering will start from the beginning. 
    // These list items will have symbols of "a.", "b.", "c.", "d.", and "e".
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"DOC");
    builder->Writeln(u"PDF");
    builder->Writeln(u"HTML");
    builder->Writeln(u"MHTML");
    builder->Writeln(u"Plain text");
    
    // Outdent the list level again.
    builder->get_ListFormat()->ListOutdent();
    builder->Writeln(u"Doing many other things!");
    
    // End the numbered list.
    builder->get_ListFormat()->RemoveNumbers();
    
    doc->Save(get_ArtifactsDir() + u"Lists.ApplyDefaultBulletsAndNumbers.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.ApplyDefaultBulletsAndNumbers.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.", 18.0, Aspose::Words::NumberStyle::Arabic, doc->get_Lists()->idx_get(1)->get_ListLevels()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0001.", 54.0, Aspose::Words::NumberStyle::LowercaseLetter, doc->get_Lists()->idx_get(1)->get_ListLevels()->idx_get(1));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\uf0b7", 18.0, Aspose::Words::NumberStyle::Bullet, doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExLists, ApplyDefaultBulletsAndNumbers)
{
    s_instance->ApplyDefaultBulletsAndNumbers();
}

} // namespace gtest_test

void ExLists::SpecifyListLevel()
{
    //ExStart
    //ExFor:ListCollection
    //ExFor:List
    //ExFor:ListFormat
    //ExFor:ListFormat.IsListItem
    //ExFor:ListFormat.ListLevelNumber
    //ExFor:ListFormat.List
    //ExFor:ListTemplate
    //ExFor:DocumentBase.Lists
    //ExFor:ListCollection.Add(ListTemplate)
    //ExSummary:Shows how to work with list levels.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    ASSERT_FALSE(builder->get_ListFormat()->get_IsListItem());
    
    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Below are two types of lists that we can create using a document builder.
    // 1 -  A numbered list:
    // Numbered lists create a logical order for their paragraphs by numbering each item.
    builder->get_ListFormat()->set_List(doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault));
    
    ASSERT_TRUE(builder->get_ListFormat()->get_IsListItem());
    
    // By setting the "ListLevelNumber" property, we can increase the list level
    // to begin a self-contained sub-list at the current list item.
    // The Microsoft Word list template called "NumberDefault" uses numbers to create list levels for the first list level.
    // Deeper list levels use letters and lowercase Roman numerals. 
    for (int32_t i = 0; i < 9; i++)
    {
        builder->get_ListFormat()->set_ListLevelNumber(i);
        builder->Writeln(System::String(u"Level ") + i);
    }
    
    // 2 -  A bulleted list:
    // This list will apply an indent and a bullet symbol ("•") before each paragraph.
    // Deeper levels of this list will use different symbols, such as "■" and "○".
    builder->get_ListFormat()->set_List(doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::BulletDefault));
    
    for (int32_t i = 0; i < 9; i++)
    {
        builder->get_ListFormat()->set_ListLevelNumber(i);
        builder->Writeln(System::String(u"Level ") + i);
    }
    
    // We can disable list formatting to not format any subsequent paragraphs as lists by un-setting the "List" flag.
    builder->get_ListFormat()->set_List(nullptr);
    
    ASSERT_FALSE(builder->get_ListFormat()->get_IsListItem());
    
    doc->Save(get_ArtifactsDir() + u"Lists.SpecifyListLevel.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.SpecifyListLevel.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.", 18.0, Aspose::Words::NumberStyle::Arabic, doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExLists, SpecifyListLevel)
{
    s_instance->SpecifyListLevel();
}

} // namespace gtest_test

void ExLists::NestedLists()
{
    //ExStart
    //ExFor:ListFormat.List
    //ExFor:ParagraphFormat.ClearFormatting
    //ExFor:ParagraphFormat.DropCapPosition
    //ExFor:ParagraphFormat.IsListItem
    //ExFor:Paragraph.IsListItem
    //ExSummary:Shows how to nest a list inside another list.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Create an outline list for the headings.
    System::SharedPtr<Aspose::Words::Lists::List> outlineList = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::OutlineNumbers);
    builder->get_ListFormat()->set_List(outlineList);
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Writeln(u"This is my Chapter 1");
    
    // Create a numbered list.
    System::SharedPtr<Aspose::Words::Lists::List> numberedList = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);
    builder->get_ListFormat()->set_List(numberedList);
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Normal);
    builder->Writeln(u"Numbered list item 1.");
    
    // Every paragraph that comprises a list will have this flag.
    ASSERT_TRUE(builder->get_CurrentParagraph()->get_IsListItem());
    ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsListItem());
    
    // Create a bulleted list.
    System::SharedPtr<Aspose::Words::Lists::List> bulletedList = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::BulletDefault);
    builder->get_ListFormat()->set_List(bulletedList);
    builder->get_ParagraphFormat()->set_LeftIndent(72);
    builder->Writeln(u"Bulleted list item 1.");
    builder->Writeln(u"Bulleted list item 2.");
    builder->get_ParagraphFormat()->ClearFormatting();
    
    // Revert to the numbered list.
    builder->get_ListFormat()->set_List(numberedList);
    builder->Writeln(u"Numbered list item 2.");
    builder->Writeln(u"Numbered list item 3.");
    
    // Revert to the outline list.
    builder->get_ListFormat()->set_List(outlineList);
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Writeln(u"This is my Chapter 2");
    
    builder->get_ParagraphFormat()->ClearFormatting();
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"Lists.NestedLists.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.NestedLists.docx");
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000)", 0.0, Aspose::Words::NumberStyle::Arabic, doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.", 18.0, Aspose::Words::NumberStyle::Arabic, doc->get_Lists()->idx_get(1)->get_ListLevels()->idx_get(0));
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\uf0b7", 18.0, Aspose::Words::NumberStyle::Bullet, doc->get_Lists()->idx_get(2)->get_ListLevels()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExLists, NestedLists)
{
    s_instance->NestedLists();
}

} // namespace gtest_test

void ExLists::CreateCustomList()
{
    //ExStart
    //ExFor:List
    //ExFor:List.ListLevels
    //ExFor:ListFormat.ListLevel
    //ExFor:ListLevelCollection
    //ExFor:ListLevelCollection.Item
    //ExFor:ListLevel
    //ExFor:ListLevel.Alignment
    //ExFor:ListLevel.Font
    //ExFor:ListLevel.NumberStyle
    //ExFor:ListLevel.StartAt
    //ExFor:ListLevel.TrailingCharacter
    //ExFor:ListLevelAlignment
    //ExFor:NumberStyle
    //ExFor:ListTrailingCharacter
    //ExFor:ListLevel.NumberFormat
    //ExFor:ListLevel.NumberPosition
    //ExFor:ListLevel.TextPosition
    //ExFor:ListLevel.TabPosition
    //ExSummary:Shows how to apply custom list formatting to paragraphs when using DocumentBuilder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Create a list from a Microsoft Word template, and customize the first two of its list levels.
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);
    
    System::SharedPtr<Aspose::Words::Lists::ListLevel> listLevel = list->get_ListLevels()->idx_get(0);
    listLevel->get_Font()->set_Color(System::Drawing::Color::get_Red());
    listLevel->get_Font()->set_Size(24);
    listLevel->set_NumberStyle(Aspose::Words::NumberStyle::OrdinalText);
    listLevel->set_StartAt(21);
    listLevel->set_NumberFormat(u"\x0000");
    
    listLevel->set_NumberPosition(-36);
    listLevel->set_TextPosition(144);
    listLevel->set_TabPosition(144);
    
    listLevel = list->get_ListLevels()->idx_get(1);
    listLevel->set_Alignment(Aspose::Words::Lists::ListLevelAlignment::Right);
    listLevel->set_NumberStyle(Aspose::Words::NumberStyle::Bullet);
    listLevel->get_Font()->set_Name(u"Wingdings");
    listLevel->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    listLevel->get_Font()->set_Size(24);
    
    // This NumberFormat value will create star-shaped bullet list symbols.
    listLevel->set_NumberFormat(u"\xf0af");
    listLevel->set_TrailingCharacter(Aspose::Words::Lists::ListTrailingCharacter::Space);
    listLevel->set_NumberPosition(144);
    
    // Create paragraphs and apply both list levels of our custom list formatting to them.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ListFormat()->set_List(list);
    builder->Writeln(u"The quick brown fox...");
    builder->Writeln(u"The quick brown fox...");
    
    builder->get_ListFormat()->ListIndent();
    builder->Writeln(u"jumped over the lazy dog.");
    builder->Writeln(u"jumped over the lazy dog.");
    
    builder->get_ListFormat()->ListOutdent();
    builder->Writeln(u"The quick brown fox...");
    
    builder->get_ListFormat()->RemoveNumbers();
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"Lists.CreateCustomList.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.CreateCustomList.docx");
    
    listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000", -36.0, Aspose::Words::NumberStyle::OrdinalText, listLevel);
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), listLevel->get_Font()->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(24.0, listLevel->get_Font()->get_Size());
    ASSERT_EQ(21, listLevel->get_StartAt());
    
    listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(1);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\xf0af", 144.0, Aspose::Words::NumberStyle::Bullet, listLevel);
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), listLevel->get_Font()->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(24.0, listLevel->get_Font()->get_Size());
    ASSERT_EQ(1, listLevel->get_StartAt());
    ASSERT_EQ(Aspose::Words::Lists::ListTrailingCharacter::Space, listLevel->get_TrailingCharacter());
}

namespace gtest_test
{

TEST_F(ExLists, CreateCustomList)
{
    s_instance->CreateCustomList();
}

} // namespace gtest_test

void ExLists::RestartNumberingUsingListCopy()
{
    //ExStart
    //ExFor:List
    //ExFor:ListCollection
    //ExFor:ListCollection.Add(ListTemplate)
    //ExFor:ListCollection.AddCopy(List)
    //ExFor:ListLevel.StartAt
    //ExFor:ListTemplate
    //ExSummary:Shows how to restart numbering in a list by copying a list.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Create a list from a Microsoft Word template, and customize its first list level.
    System::SharedPtr<Aspose::Words::Lists::List> list1 = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberArabicParenthesis);
    list1->get_ListLevels()->idx_get(0)->get_Font()->set_Color(System::Drawing::Color::get_Red());
    list1->get_ListLevels()->idx_get(0)->set_Alignment(Aspose::Words::Lists::ListLevelAlignment::Right);
    
    // Apply our list to some paragraphs.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"List 1 starts below:");
    builder->get_ListFormat()->set_List(list1);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->RemoveNumbers();
    
    // We can add a copy of an existing list to the document's list collection
    // to create a similar list without making changes to the original.
    System::SharedPtr<Aspose::Words::Lists::List> list2 = doc->get_Lists()->AddCopy(list1);
    list2->get_ListLevels()->idx_get(0)->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    list2->get_ListLevels()->idx_get(0)->set_StartAt(10);
    
    // Apply the second list to new paragraphs.
    builder->Writeln(u"List 2 starts below:");
    builder->get_ListFormat()->set_List(list2);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->RemoveNumbers();
    
    doc->Save(get_ArtifactsDir() + u"Lists.RestartNumberingUsingListCopy.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.RestartNumberingUsingListCopy.docx");
    
    list1 = doc->get_Lists()->idx_get(0);
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000)", 18.0, Aspose::Words::NumberStyle::Arabic, list1->get_ListLevels()->idx_get(0));
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), list1->get_ListLevels()->idx_get(0)->get_Font()->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(10.0, list1->get_ListLevels()->idx_get(0)->get_Font()->get_Size());
    ASSERT_EQ(1, list1->get_ListLevels()->idx_get(0)->get_StartAt());
    
    list2 = doc->get_Lists()->idx_get(1);
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000)", 18.0, Aspose::Words::NumberStyle::Arabic, list2->get_ListLevels()->idx_get(0));
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), list2->get_ListLevels()->idx_get(0)->get_Font()->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(10.0, list2->get_ListLevels()->idx_get(0)->get_Font()->get_Size());
    ASSERT_EQ(10, list2->get_ListLevels()->idx_get(0)->get_StartAt());
}

namespace gtest_test
{

TEST_F(ExLists, RestartNumberingUsingListCopy)
{
    s_instance->RestartNumberingUsingListCopy();
}

} // namespace gtest_test

void ExLists::CreateAndUseListStyle()
{
    //ExStart
    //ExFor:StyleCollection.Add(StyleType,String)
    //ExFor:Style.List
    //ExFor:StyleType
    //ExFor:List.IsListStyleDefinition
    //ExFor:List.IsListStyleReference
    //ExFor:List.IsMultiLevel
    //ExFor:List.Style
    //ExFor:ListLevelCollection
    //ExFor:ListLevelCollection.Count
    //ExFor:ListLevelCollection.Item
    //ExFor:ListCollection.Add(Style)
    //ExSummary:Shows how to create a list style and use it in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // We can contain an entire List object within a style.
    System::SharedPtr<Aspose::Words::Style> listStyle = doc->get_Styles()->Add(Aspose::Words::StyleType::List, u"MyListStyle");
    
    System::SharedPtr<Aspose::Words::Lists::List> list1 = listStyle->get_List();
    
    ASSERT_TRUE(list1->get_IsListStyleDefinition());
    ASSERT_FALSE(list1->get_IsListStyleReference());
    ASSERT_TRUE(list1->get_IsMultiLevel());
    ASPOSE_ASSERT_EQ(listStyle, list1->get_Style());
    
    // Change the appearance of all list levels in our list.
    for (auto&& level : list1->get_ListLevels())
    {
        level->get_Font()->set_Name(u"Verdana");
        level->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        level->get_Font()->set_Bold(true);
    }
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Using list style first time:");
    
    // Create another list from a list within a style.
    System::SharedPtr<Aspose::Words::Lists::List> list2 = doc->get_Lists()->Add(listStyle);
    
    ASSERT_FALSE(list2->get_IsListStyleDefinition());
    ASSERT_TRUE(list2->get_IsListStyleReference());
    ASPOSE_ASSERT_EQ(listStyle, list2->get_Style());
    
    // Add some list items that our list will format.
    builder->get_ListFormat()->set_List(list2);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->RemoveNumbers();
    
    builder->Writeln(u"Using list style second time:");
    
    // Create and apply another list based on the list style.
    System::SharedPtr<Aspose::Words::Lists::List> list3 = doc->get_Lists()->Add(listStyle);
    builder->get_ListFormat()->set_List(list3);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->RemoveNumbers();
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"Lists.CreateAndUseListStyle.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.CreateAndUseListStyle.docx");
    
    list1 = doc->get_Lists()->idx_get(0);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.", 18.0, Aspose::Words::NumberStyle::Arabic, list1->get_ListLevels()->idx_get(0));
    ASSERT_TRUE(list1->get_IsListStyleDefinition());
    ASSERT_FALSE(list1->get_IsListStyleReference());
    ASSERT_TRUE(list1->get_IsMultiLevel());
    ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), list1->get_ListLevels()->idx_get(0)->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(u"Verdana", list1->get_ListLevels()->idx_get(0)->get_Font()->get_Name());
    ASSERT_TRUE(list1->get_ListLevels()->idx_get(0)->get_Font()->get_Bold());
    
    list2 = doc->get_Lists()->idx_get(1);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.", 18.0, Aspose::Words::NumberStyle::Arabic, list2->get_ListLevels()->idx_get(0));
    ASSERT_FALSE(list2->get_IsListStyleDefinition());
    ASSERT_TRUE(list2->get_IsListStyleReference());
    ASSERT_TRUE(list2->get_IsMultiLevel());
    
    list3 = doc->get_Lists()->idx_get(2);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"\u0000.", 18.0, Aspose::Words::NumberStyle::Arabic, list3->get_ListLevels()->idx_get(0));
    ASSERT_FALSE(list3->get_IsListStyleDefinition());
    ASSERT_TRUE(list3->get_IsListStyleReference());
    ASSERT_TRUE(list3->get_IsMultiLevel());
}

namespace gtest_test
{

TEST_F(ExLists, CreateAndUseListStyle)
{
    s_instance->CreateAndUseListStyle();
}

} // namespace gtest_test

void ExLists::DetectBulletedParagraphs()
{
    //ExStart
    //ExFor:Paragraph.ListFormat
    //ExFor:ListFormat.IsListItem
    //ExFor:CompositeNode.GetText
    //ExFor:List.ListId
    //ExSummary:Shows how to output all paragraphs in a document that are list items.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ListFormat()->ApplyNumberDefault();
    builder->Writeln(u"Numbered list item 1");
    builder->Writeln(u"Numbered list item 2");
    builder->Writeln(u"Numbered list item 3");
    builder->get_ListFormat()->RemoveNumbers();
    
    builder->get_ListFormat()->ApplyBulletDefault();
    builder->Writeln(u"Bulleted list item 1");
    builder->Writeln(u"Bulleted list item 2");
    builder->Writeln(u"Bulleted list item 3");
    builder->get_ListFormat()->RemoveNumbers();
    
    System::SharedPtr<Aspose::Words::NodeCollection> paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    
    for (auto&& para : paras->LINQ_OfType<System::SharedPtr<Aspose::Words::Paragraph> >()->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Paragraph>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Paragraph> p)>>([](System::SharedPtr<Aspose::Words::Paragraph> p) -> bool
    {
        return p->get_ListFormat()->get_IsListItem();
    })))->LINQ_ToList())
    {
        std::cout << System::String::Format(u"This paragraph belongs to list ID# {0}, number style \"{1}\"", para->get_ListFormat()->get_List()->get_ListId(), para->get_ListFormat()->get_ListLevel()->get_NumberStyle()) << std::endl;
        std::cout << System::String::Format(u"\t\"{0}\"", para->GetText().Trim()) << std::endl;
    }
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    
    ASSERT_EQ(6, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
}

namespace gtest_test
{

TEST_F(ExLists, DetectBulletedParagraphs)
{
    s_instance->DetectBulletedParagraphs();
}

} // namespace gtest_test

void ExLists::RemoveBulletsFromParagraphs()
{
    //ExStart
    //ExFor:ListFormat.RemoveNumbers
    //ExSummary:Shows how to remove list formatting from all paragraphs in the main text of a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ListFormat()->ApplyNumberDefault();
    builder->Writeln(u"Numbered list item 1");
    builder->Writeln(u"Numbered list item 2");
    builder->Writeln(u"Numbered list item 3");
    builder->get_ListFormat()->RemoveNumbers();
    
    System::SharedPtr<Aspose::Words::NodeCollection> paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    
    for (auto&& paragraph : System::IterateOver<Aspose::Words::Paragraph>(paras))
    {
        paragraph->get_ListFormat()->RemoveNumbers();
    }
    
    ASSERT_EQ(0, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLists, RemoveBulletsFromParagraphs)
{
    s_instance->RemoveBulletsFromParagraphs();
}

} // namespace gtest_test

void ExLists::ApplyExistingListToParagraphs()
{
    //ExStart
    //ExFor:ListCollection.Item(Int32)
    //ExSummary:Shows how to apply list formatting of an existing list to a collection of paragraphs.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Paragraph 1");
    builder->Writeln(u"Paragraph 2");
    builder->Write(u"Paragraph 3");
    
    System::SharedPtr<Aspose::Words::NodeCollection> paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    
    ASSERT_EQ(0, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    
    doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->idx_get(0);
    
    for (auto&& paragraph : System::IterateOver(paras->LINQ_OfType<System::SharedPtr<Aspose::Words::Paragraph> >()))
    {
        paragraph->get_ListFormat()->set_List(list);
        paragraph->get_ListFormat()->set_ListLevelNumber(2);
    }
    
    ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    
    ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_ListLevelNumber() == 2;
    }))));
}

namespace gtest_test
{

TEST_F(ExLists, ApplyExistingListToParagraphs)
{
    s_instance->ApplyExistingListToParagraphs();
}

} // namespace gtest_test

void ExLists::ApplyNewListToParagraphs()
{
    //ExStart
    //ExFor:ListCollection.Add(ListTemplate)
    //ExSummary:Shows how to create a list by applying a new list format to a collection of paragraphs.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Paragraph 1");
    builder->Writeln(u"Paragraph 2");
    builder->Write(u"Paragraph 3");
    
    System::SharedPtr<Aspose::Words::NodeCollection> paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    
    ASSERT_EQ(0, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberUppercaseLetterDot);
    
    for (auto&& paragraph : System::IterateOver(paras->LINQ_OfType<System::SharedPtr<Aspose::Words::Paragraph> >()))
    {
        paragraph->get_ListFormat()->set_List(list);
        paragraph->get_ListFormat()->set_ListLevelNumber(1);
    }
    
    ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    
    ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_IsListItem();
    }))));
    ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Node>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Node> n)>>([](System::SharedPtr<Aspose::Words::Node> n) -> bool
    {
        return (System::ExplicitCast<Aspose::Words::Paragraph>(n))->get_ListFormat()->get_ListLevelNumber() == 1;
    }))));
}

namespace gtest_test
{

TEST_F(ExLists, ApplyNewListToParagraphs)
{
    s_instance->ApplyNewListToParagraphs();
}

} // namespace gtest_test

void ExLists::OutlineHeadingTemplates()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::OutlineHeadingsArticleSection);
    AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Article Section\"");
    
    list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::OutlineHeadingsLegal);
    AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Legal\"");
    
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::OutlineHeadingsNumbers);
    AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Numbers\"");
    
    list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::OutlineHeadingsChapter);
    AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Chapters\"");
    
    doc->Save(get_ArtifactsDir() + u"Lists.OutlineHeadingTemplates.docx");
    TestOutlineHeadingTemplates(System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.OutlineHeadingTemplates.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExLists, OutlineHeadingTemplates)
{
    s_instance->OutlineHeadingTemplates();
}

} // namespace gtest_test

void ExLists::PrintOutAllLists()
{
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    
    for (auto&& srcList : srcDoc->get_Lists())
    {
        System::SharedPtr<Aspose::Words::Lists::List> dstList = dstDoc->get_Lists()->AddCopy(srcList);
        AddListSample(builder, dstList);
    }
    
    dstDoc->Save(get_ArtifactsDir() + u"Lists.PrintOutAllLists.docx");
    TestPrintOutAllLists(srcDoc, System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.PrintOutAllLists.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExLists, PrintOutAllLists)
{
    s_instance->PrintOutAllLists();
}

} // namespace gtest_test

void ExLists::ListDocument()
{
    //ExStart
    //ExFor:ListCollection.Document
    //ExFor:ListCollection.Count
    //ExFor:ListCollection.Item(Int32)
    //ExFor:ListCollection.GetListByListId
    //ExFor:List.Document
    //ExFor:List.ListId
    //ExSummary:Shows how to verify owner document properties of lists.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Lists::ListCollection> lists = doc->get_Lists();
    ASPOSE_ASSERT_EQ(doc, lists->get_Document());
    
    System::SharedPtr<Aspose::Words::Lists::List> list = lists->Add(Aspose::Words::Lists::ListTemplate::BulletDefault);
    ASPOSE_ASSERT_EQ(doc, list->get_Document());
    
    std::cout << (System::String(u"Current list count: ") + lists->get_Count()) << std::endl;
    std::cout << (System::String(u"Is the first document list: ") + (System::ObjectExt::Equals(lists->idx_get(0), list))) << std::endl;
    std::cout << (System::String(u"ListId: ") + list->get_ListId()) << std::endl;
    std::cout << (System::String(u"List is the same by ListId: ") + (System::ObjectExt::Equals(lists->GetListByListId(1), list))) << std::endl;
    //ExEnd
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    lists = doc->get_Lists();
    
    ASPOSE_ASSERT_EQ(doc, lists->get_Document());
    ASSERT_EQ(1, lists->get_Count());
    ASSERT_EQ(1, lists->idx_get(0)->get_ListId());
    ASPOSE_ASSERT_EQ(lists->idx_get(0), lists->GetListByListId(1));
}

namespace gtest_test
{

TEST_F(ExLists, ListDocument)
{
    s_instance->ListDocument();
}

} // namespace gtest_test

void ExLists::CreateListRestartAfterHigher()
{
    //ExStart
    //ExFor:ListLevel.NumberStyle
    //ExFor:ListLevel.NumberFormat
    //ExFor:ListLevel.IsLegal
    //ExFor:ListLevel.RestartAfterLevel
    //ExFor:ListLevel.LinkedStyle
    //ExSummary:Shows advances ways of customizing list labels.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);
    
    // Level 1 labels will be formatted according to the "Heading 1" paragraph style and will have a prefix.
    // These will look like "Appendix A", "Appendix B"...
    list->get_ListLevels()->idx_get(0)->set_NumberFormat(u"Appendix \x0000");
    list->get_ListLevels()->idx_get(0)->set_NumberStyle(Aspose::Words::NumberStyle::UppercaseLetter);
    list->get_ListLevels()->idx_get(0)->set_LinkedStyle(doc->get_Styles()->idx_get(u"Heading 1"));
    
    // Level 2 labels will display the current numbers of the first and the second list levels and have leading zeroes.
    // If the first list level is at 1, then the list labels from these will look like "Section (1.01)", "Section (1.02)"...
    list->get_ListLevels()->idx_get(1)->set_NumberFormat(u"Section (\x0000" u".\x0001" u")");
    list->get_ListLevels()->idx_get(1)->set_NumberStyle(Aspose::Words::NumberStyle::LeadingZero);
    
    // Note that the higher-level uses UppercaseLetter numbering.
    // We can set the "IsLegal" property to use Arabic numbers for the higher list levels.
    list->get_ListLevels()->idx_get(1)->set_IsLegal(true);
    list->get_ListLevels()->idx_get(1)->set_RestartAfterLevel(0);
    
    // Level 3 labels will be upper case Roman numerals with a prefix and a suffix and will restart at each List level 1 item.
    // These list labels will look like "-I-", "-II-"...
    list->get_ListLevels()->idx_get(2)->set_NumberFormat(u"-\x0002" u"-");
    list->get_ListLevels()->idx_get(2)->set_NumberStyle(Aspose::Words::NumberStyle::UppercaseRoman);
    list->get_ListLevels()->idx_get(2)->set_RestartAfterLevel(1);
    
    // Make labels of all list levels bold.
    for (auto&& level : list->get_ListLevels())
    {
        level->get_Font()->set_Bold(true);
    }
    
    // Apply list formatting to the current paragraph.
    builder->get_ListFormat()->set_List(list);
    
    // Create list items that will display all three of our list levels.
    for (int32_t n = 0; n < 2; n++)
    {
        for (int32_t i = 0; i < 3; i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(System::String(u"Level ") + i);
        }
    }
    
    builder->get_ListFormat()->RemoveNumbers();
    
    doc->Save(get_ArtifactsDir() + u"Lists.CreateListRestartAfterHigher.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.CreateListRestartAfterHigher.docx");
    
    System::SharedPtr<Aspose::Words::Lists::ListLevel> listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"Appendix \u0000", 18.0, Aspose::Words::NumberStyle::UppercaseLetter, listLevel);
    ASSERT_FALSE(listLevel->get_IsLegal());
    ASSERT_EQ(-1, listLevel->get_RestartAfterLevel());
    ASSERT_EQ(u"Heading 1", listLevel->get_LinkedStyle()->get_Name());
    
    listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(1);
    
    Aspose::Words::ApiExamples::TestUtil::VerifyListLevel(u"Section (\u0000.\u0001)", 54.0, Aspose::Words::NumberStyle::LeadingZero, listLevel);
    ASSERT_TRUE(listLevel->get_IsLegal());
    ASSERT_EQ(0, listLevel->get_RestartAfterLevel());
    ASSERT_TRUE(System::TestTools::IsNull(listLevel->get_LinkedStyle()));
}

namespace gtest_test
{

TEST_F(ExLists, CreateListRestartAfterHigher)
{
    s_instance->CreateListRestartAfterHigher();
}

} // namespace gtest_test

void ExLists::GetListLabels()
{
    //ExStart
    //ExFor:Document.UpdateListLabels()
    //ExFor:Node.ToString(SaveFormat)
    //ExFor:ListLabel
    //ExFor:Paragraph.ListLabel
    //ExFor:ListLabel.LabelValue
    //ExFor:ListLabel.LabelString
    //ExSummary:Shows how to extract the list labels of all paragraphs that are list items.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    doc->UpdateListLabels();
    
    System::SharedPtr<Aspose::Words::NodeCollection> paras = doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true);
    
    // Find if we have the paragraph list. In our document, our list uses plain Arabic numbers,
    // which start at three and ends at six.
    for (auto&& paragraph : paras->LINQ_OfType<System::SharedPtr<Aspose::Words::Paragraph> >()->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Paragraph>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Paragraph> p)>>([](System::SharedPtr<Aspose::Words::Paragraph> p) -> bool
    {
        return p->get_ListFormat()->get_IsListItem();
    })))->LINQ_ToList())
    {
        std::cout << System::String::Format(u"List item paragraph #{0}", paras->IndexOf(paragraph)) << std::endl;
        
        // This is the text we get when getting when we output this node to text format.
        // This text output will omit list labels. Trim any paragraph formatting characters. 
        System::String paragraphText = paragraph->ToString(Aspose::Words::SaveFormat::Text).Trim();
        std::cout << System::String::Format(u"\tExported Text: {0}", paragraphText) << std::endl;
        
        System::SharedPtr<Aspose::Words::Lists::ListLabel> label = paragraph->get_ListLabel();
        
        // This gets the position of the paragraph in the current level of the list. If we have a list with multiple levels,
        // this will tell us what position it is on that level.
        std::cout << System::String::Format(u"\tNumerical Id: {0}", label->get_LabelValue()) << std::endl;
        
        // Combine them together to include the list label with the text in the output.
        std::cout << System::String::Format(u"\tList label combined with text: {0} {1}", label->get_LabelString(), paragraphText) << std::endl;
    }
    //ExEnd
    
    ASSERT_EQ(10, paras->LINQ_OfType<System::SharedPtr<Aspose::Words::Paragraph> >()->LINQ_Count(static_cast<System::Func<System::SharedPtr<Aspose::Words::Paragraph>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Paragraph> p)>>([](System::SharedPtr<Aspose::Words::Paragraph> p) -> bool
    {
        return p->get_ListFormat()->get_IsListItem();
    }))));
}

namespace gtest_test
{

TEST_F(ExLists, GetListLabels)
{
    s_instance->GetListLabels();
}

} // namespace gtest_test

void ExLists::CreatePictureBullet()
{
    //ExStart
    //ExFor:ListLevel.CreatePictureBullet
    //ExFor:ListLevel.DeletePictureBullet
    //ExSummary:Shows how to set a custom image icon for list item labels.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::BulletCircle);
    
    // Create a picture bullet for the current list level, and set an image from a local file system
    // as the icon that the bullets for this list level will display.
    list->get_ListLevels()->idx_get(0)->CreatePictureBullet();
    list->get_ListLevels()->idx_get(0)->get_ImageData()->SetImage(get_ImageDir() + u"Logo icon.ico");
    
    ASSERT_TRUE(list->get_ListLevels()->idx_get(0)->get_ImageData()->get_HasImage());
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_ListFormat()->set_List(list);
    builder->Writeln(u"Hello world!");
    builder->Write(u"Hello again!");
    
    doc->Save(get_ArtifactsDir() + u"Lists.CreatePictureBullet.docx");
    
    list->get_ListLevels()->idx_get(0)->DeletePictureBullet();
    
    ASSERT_TRUE(System::TestTools::IsNull(list->get_ListLevels()->idx_get(0)->get_ImageData()));
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Lists.CreatePictureBullet.docx");
    
    ASSERT_TRUE(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData()->get_HasImage());
}

namespace gtest_test
{

TEST_F(ExLists, IgnoreOnJenkins_CreatePictureBullet)
{
    RecordProperty("category", "IgnoreOnJenkins");
    
    s_instance->CreatePictureBullet();
}

} // namespace gtest_test

void ExLists::GetCustomNumberStyleFormat()
{
    //ExStart
    //ExFor:ListLevel.CustomNumberStyleFormat
    //ExFor:ListLevel.GetEffectiveValue(Int32, NumberStyle, String)
    //ExSummary:Shows how to get the format for a list with the custom number style.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List with leading zero.docx");
    
    System::SharedPtr<Aspose::Words::Lists::ListLevel> listLevel = doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ListFormat()->get_ListLevel();
    
    System::String customNumberStyleFormat = System::String::Empty;
    
    if (listLevel->get_NumberStyle() == Aspose::Words::NumberStyle::Custom)
    {
        customNumberStyleFormat = listLevel->get_CustomNumberStyleFormat();
    }
    
    ASSERT_EQ(u"001, 002, 003, ...", customNumberStyleFormat);
    
    // We can get value for the specified index of the list item.
    ASSERT_EQ(u"iv", Aspose::Words::Lists::ListLevel::GetEffectiveValue(4, Aspose::Words::NumberStyle::LowercaseRoman, nullptr));
    ASSERT_EQ(u"005", Aspose::Words::Lists::ListLevel::GetEffectiveValue(5, Aspose::Words::NumberStyle::Custom, customNumberStyleFormat));
    //ExEnd
    
    ASSERT_THROW(static_cast<std::function<void()>>([&customNumberStyleFormat]() -> void
    {
        Aspose::Words::Lists::ListLevel::GetEffectiveValue(5, Aspose::Words::NumberStyle::LowercaseRoman, customNumberStyleFormat);
    })(), System::ArgumentException);
    ASSERT_THROW(static_cast<std::function<void()>>([]() -> void
    {
        Aspose::Words::Lists::ListLevel::GetEffectiveValue(5, Aspose::Words::NumberStyle::Custom, nullptr);
    })(), System::ArgumentException);
    ASSERT_THROW(static_cast<std::function<void()>>([]() -> void
    {
        Aspose::Words::Lists::ListLevel::GetEffectiveValue(5, Aspose::Words::NumberStyle::Custom, u"....");
    })(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExLists, GetCustomNumberStyleFormat)
{
    s_instance->GetCustomNumberStyleFormat();
}

} // namespace gtest_test

void ExLists::HasSameTemplate()
{
    //ExStart
    //ExFor:List.HasSameTemplate(List)
    //ExSummary:Shows how to define lists with the same ListDefId.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Different lists.docx");
    
    ASSERT_TRUE(doc->get_Lists()->idx_get(0)->HasSameTemplate(doc->get_Lists()->idx_get(1)));
    ASSERT_FALSE(doc->get_Lists()->idx_get(1)->HasSameTemplate(doc->get_Lists()->idx_get(2)));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLists, HasSameTemplate)
{
    s_instance->HasSameTemplate();
}

} // namespace gtest_test

void ExLists::SetCustomNumberStyleFormat()
{
    //ExStart:SetCustomNumberStyleFormat
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:ListLevel.CustomNumberStyleFormat
    //ExSummary:Shows how to set customer number style format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"List with leading zero.docx");
    
    doc->UpdateListLabels();
    
    System::SharedPtr<Aspose::Words::ParagraphCollection> paras = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    ASSERT_EQ(u"001.", paras->idx_get(0)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"0001.", paras->idx_get(1)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"0002.", paras->idx_get(2)->get_ListLabel()->get_LabelString());
    
    paras->idx_get(1)->get_ListFormat()->get_ListLevel()->set_CustomNumberStyleFormat(u"001, 002, 003, ...");
    
    doc->UpdateListLabels();
    
    ASSERT_EQ(u"001.", paras->idx_get(0)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"001.", paras->idx_get(1)->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"002.", paras->idx_get(2)->get_ListLabel()->get_LabelString());
    //ExEnd:SetCustomNumberStyleFormat
}

namespace gtest_test
{

TEST_F(ExLists, SetCustomNumberStyleFormat)
{
    s_instance->SetCustomNumberStyleFormat();
}

} // namespace gtest_test

void ExLists::AddSingleLevelList()
{
    //ExStart:AddSingleLevelList
    //GistId:95fdae949cefbf2ce485acc95cccc495
    //ExFor:ListCollection.AddSingleLevelList(ListTemplate)
    //ExSummary:Shows how to create a new single level list based on the predefined template.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Lists::ListCollection> listCollection = doc->get_Lists();
    
    // Creates the bulleted list from BulletCircle template.
    System::SharedPtr<Aspose::Words::Lists::List> bulletedList = listCollection->AddSingleLevelList(Aspose::Words::Lists::ListTemplate::BulletCircle);
    
    // Writes the bulleted list to the resulting document.
    builder->Writeln(u"Bulleted list starts below:");
    builder->get_ListFormat()->set_List(bulletedList);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    builder->get_ListFormat()->RemoveNumbers();
    
    // Creates the numbered list from NumberUppercaseLetterDot template.
    System::SharedPtr<Aspose::Words::Lists::List> numberedList = listCollection->AddSingleLevelList(Aspose::Words::Lists::ListTemplate::NumberUppercaseLetterDot);
    
    // Writes the numbered list to the resulting document.
    builder->Writeln(u"Numbered list starts below:");
    builder->get_ListFormat()->set_List(numberedList);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");
    
    doc->Save(get_ArtifactsDir() + u"Lists.AddSingleLevelList.docx");
    //ExEnd:AddSingleLevelList
}

namespace gtest_test
{

TEST_F(ExLists, AddSingleLevelList)
{
    s_instance->AddSingleLevelList();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
