#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelAlignment.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTrailingCharacter.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;

namespace ApiExamples {

class ExLists : public ApiExampleBase
{
public:
    void ApplyDefaultBulletsAndNumbers()
    {
        //ExStart
        //ExFor:DocumentBuilder.ListFormat
        //ExFor:ListFormat.ApplyNumberDefault
        //ExFor:ListFormat.ApplyBulletDefault
        //ExFor:ListFormat.ListIndent
        //ExFor:ListFormat.ListOutdent
        //ExFor:ListFormat.RemoveNumbers
        //ExFor:ListFormat.ListLevelNumber
        //ExSummary:Shows how to apply default bulleted or numbered list formatting to paragraphs when using DocumentBuilder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Aspose.Words allows:");
        builder->Writeln();

        // Start a numbered list with default formatting
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Opening documents from different formats:");

        ASSERT_EQ(0, builder->get_ListFormat()->get_ListLevelNumber());

        // Go to second list level, add more text
        builder->get_ListFormat()->ListIndent();

        ASSERT_EQ(1, builder->get_ListFormat()->get_ListLevelNumber());

        builder->Writeln(u"DOC");
        builder->Writeln(u"PDF");
        builder->Writeln(u"HTML");

        // Outdent to the first list level
        builder->get_ListFormat()->ListOutdent();

        ASSERT_EQ(0, builder->get_ListFormat()->get_ListLevelNumber());

        builder->Writeln(u"Processing documents");
        builder->Writeln(u"Saving documents in different formats:");

        // Indent the list level again
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"DOC");
        builder->Writeln(u"PDF");
        builder->Writeln(u"HTML");
        builder->Writeln(u"MHTML");
        builder->Writeln(u"Plain text");

        // Outdent the list level again
        builder->get_ListFormat()->ListOutdent();
        builder->Writeln(u"Doing many other things!");

        // End the numbered list
        builder->get_ListFormat()->RemoveNumbers();
        builder->Writeln();

        builder->Writeln(u"Aspose.Words main advantages are:");
        builder->Writeln();

        // Start a bulleted list with default formatting
        builder->get_ListFormat()->ApplyBulletDefault();
        builder->Writeln(u"Great performance");
        builder->Writeln(u"High reliability");
        builder->Writeln(u"Quality code and working");
        builder->Writeln(u"Wide variety of features");
        builder->Writeln(u"Easy to understand API");

        // End the bulleted list
        builder->get_ListFormat()->RemoveNumbers();

        doc->Save(ArtifactsDir + u"Lists.ApplyDefaultBulletsAndNumbers.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.ApplyDefaultBulletsAndNumbers.docx");

        TestUtil::VerifyListLevel(u"\u0000.", 18.0, NumberStyle::Arabic, doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0));
        TestUtil::VerifyListLevel(u"\u0001.", 54.0, NumberStyle::LowercaseLetter, doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(1));
        TestUtil::VerifyListLevel(u"\uf0b7", 18.0, NumberStyle::Bullet, doc->get_Lists()->idx_get(1)->get_ListLevels()->idx_get(0));
    }

    void SpecifyListLevel()
    {
        //ExStart
        //ExFor:ListCollection
        //ExFor:List
        //ExFor:ListFormat
        //ExFor:ListFormat.ListLevelNumber
        //ExFor:ListFormat.List
        //ExFor:ListTemplate
        //ExFor:DocumentBase.Lists
        //ExFor:ListCollection.Add(ListTemplate)
        //ExSummary:Shows how to specify list level number when building a list using DocumentBuilder.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a numbered list based on one of the Microsoft Word list templates and
        // apply it to the current paragraph in the document builder
        builder->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::NumberArabicDot));

        // Insert text at each of the 9 indent levels
        for (int i = 0; i < 9; i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(String(u"Level ") + i);
        }

        // Create a bulleted list based on one of the Microsoft Word list templates
        // and apply it to the current paragraph in the document builder
        builder->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::BulletDiamonds));

        for (int i = 0; i < 9; i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(String(u"Level ") + i);
        }

        // This is a way to stop list formatting
        builder->get_ListFormat()->set_List(nullptr);

        doc->Save(ArtifactsDir + u"Lists.SpecifyListLevel.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.SpecifyListLevel.docx");

        TestUtil::VerifyListLevel(u"\u0000.", 18.0, NumberStyle::Arabic, doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0));
    }

    void NestedLists()
    {
        //ExStart
        //ExFor:ListFormat.List
        //ExFor:ParagraphFormat.ClearFormatting
        //ExFor:ParagraphFormat.DropCapPosition
        //ExFor:ParagraphFormat.IsListItem
        //ExFor:Paragraph.IsListItem
        //ExSummary:Shows how to start a numbered list, add a bulleted list inside it, then return to the numbered list.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create an outline list for the headings
        SharedPtr<List> outlineList = doc->get_Lists()->Add(ListTemplate::OutlineNumbers);
        builder->get_ListFormat()->set_List(outlineList);
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Writeln(u"This is my Chapter 1");

        // Create a numbered list
        SharedPtr<List> numberedList = doc->get_Lists()->Add(ListTemplate::NumberDefault);
        builder->get_ListFormat()->set_List(numberedList);
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Normal);
        builder->Writeln(u"Numbered list item 1.");

        // Every paragraph that comprises a list will have this flag
        ASSERT_TRUE(builder->get_CurrentParagraph()->get_IsListItem());
        ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsListItem());

        // Create a bulleted list
        SharedPtr<List> bulletedList = doc->get_Lists()->Add(ListTemplate::BulletDefault);
        builder->get_ListFormat()->set_List(bulletedList);
        builder->get_ParagraphFormat()->set_LeftIndent(72);
        builder->Writeln(u"Bulleted list item 1.");
        builder->Writeln(u"Bulleted list item 2.");
        builder->get_ParagraphFormat()->ClearFormatting();

        // Revert to the numbered list
        builder->get_ListFormat()->set_List(numberedList);
        builder->Writeln(u"Numbered list item 2.");
        builder->Writeln(u"Numbered list item 3.");

        // Revert to the outline list
        builder->get_ListFormat()->set_List(outlineList);
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);
        builder->Writeln(u"This is my Chapter 2");

        builder->get_ParagraphFormat()->ClearFormatting();

        builder->get_Document()->Save(ArtifactsDir + u"Lists.NestedLists.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.NestedLists.docx");

        TestUtil::VerifyListLevel(u"\u0000)", 0.0, NumberStyle::Arabic, doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0));
        TestUtil::VerifyListLevel(u"\u0000.", 18.0, NumberStyle::Arabic, doc->get_Lists()->idx_get(1)->get_ListLevels()->idx_get(0));
        TestUtil::VerifyListLevel(u"\uf0b7", 18.0, NumberStyle::Bullet, doc->get_Lists()->idx_get(2)->get_ListLevels()->idx_get(0));
    }

    void CreateCustomList()
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
        auto doc = MakeObject<Document>();

        // Create a list based on one of the Microsoft Word list templates
        SharedPtr<List> list = doc->get_Lists()->Add(ListTemplate::NumberDefault);

        // Completely customize one list level
        SharedPtr<ListLevel> listLevel = list->get_ListLevels()->idx_get(0);
        listLevel->get_Font()->set_Color(System::Drawing::Color::get_Red());
        listLevel->get_Font()->set_Size(24);
        listLevel->set_NumberStyle(NumberStyle::OrdinalText);
        listLevel->set_StartAt(21);
        listLevel->set_NumberFormat(u"\x0000");

        listLevel->set_NumberPosition(-36);
        listLevel->set_TextPosition(144);
        listLevel->set_TabPosition(144);

        // Customize another list level
        listLevel = list->get_ListLevels()->idx_get(1);
        listLevel->set_Alignment(ListLevelAlignment::Right);
        listLevel->set_NumberStyle(NumberStyle::Bullet);
        listLevel->get_Font()->set_Name(u"Wingdings");
        listLevel->get_Font()->set_Color(System::Drawing::Color::get_Blue());
        listLevel->get_Font()->set_Size(24);
        listLevel->set_NumberFormat(u"\xf0af");
        // A bullet that looks like a star
        listLevel->set_TrailingCharacter(ListTrailingCharacter::Space);
        listLevel->set_NumberPosition(144);

        // Now add some text that uses the list that we created
        // It does not matter when to customize the list - before or after adding the paragraphs
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ListFormat()->set_List(list);
        builder->Writeln(u"The quick brown fox...");
        builder->Writeln(u"The quick brown fox...");

        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"jumped over the lazy dog.");
        builder->Writeln(u"jumped over the lazy dog.");

        builder->get_ListFormat()->ListOutdent();
        builder->Writeln(u"The quick brown fox...");

        builder->get_ListFormat()->RemoveNumbers();

        builder->get_Document()->Save(ArtifactsDir + u"Lists.CreateCustomList.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.CreateCustomList.docx");

        listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0);

        TestUtil::VerifyListLevel(u"\u0000", -36.0, NumberStyle::OrdinalText, listLevel);
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), listLevel->get_Font()->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(24.0, listLevel->get_Font()->get_Size());
        ASSERT_EQ(21, listLevel->get_StartAt());

        listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(1);

        TestUtil::VerifyListLevel(u"\xf0af", 144.0, NumberStyle::Bullet, listLevel);
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), listLevel->get_Font()->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(24.0, listLevel->get_Font()->get_Size());
        ASSERT_EQ(1, listLevel->get_StartAt());
        ASSERT_EQ(ListTrailingCharacter::Space, listLevel->get_TrailingCharacter());
    }

    void RestartNumberingUsingListCopy()
    {
        //ExStart
        //ExFor:List
        //ExFor:ListCollection
        //ExFor:ListCollection.Add(ListTemplate)
        //ExFor:ListCollection.AddCopy(List)
        //ExFor:ListLevel.StartAt
        //ExFor:ListTemplate
        //ExSummary:Shows how to restart numbering in a list by copying a list.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list based on a template
        SharedPtr<List> list1 = doc->get_Lists()->Add(ListTemplate::NumberArabicParenthesis);
        // Modify the formatting of the list
        list1->get_ListLevels()->idx_get(0)->get_Font()->set_Color(System::Drawing::Color::get_Red());
        list1->get_ListLevels()->idx_get(0)->set_Alignment(ListLevelAlignment::Right);

        builder->Writeln(u"List 1 starts below:");
        // Use the first list in the document for a while
        builder->get_ListFormat()->set_List(list1);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        // Now I want to reuse the first list, but need to restart numbering
        // This should be done by creating a copy of the original list formatting
        SharedPtr<List> list2 = doc->get_Lists()->AddCopy(list1);

        // We can modify the new list in any way. Including setting new start number
        list2->get_ListLevels()->idx_get(0)->set_StartAt(10);

        // Use the second list in the document
        builder->Writeln(u"List 2 starts below:");
        builder->get_ListFormat()->set_List(list2);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        doc->Save(ArtifactsDir + u"Lists.RestartNumberingUsingListCopy.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.RestartNumberingUsingListCopy.docx");

        list1 = doc->get_Lists()->idx_get(0);
        TestUtil::VerifyListLevel(u"\u0000)", 18.0, NumberStyle::Arabic, list1->get_ListLevels()->idx_get(0));
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), list1->get_ListLevels()->idx_get(0)->get_Font()->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(10.0, list1->get_ListLevels()->idx_get(0)->get_Font()->get_Size());
        ASSERT_EQ(1, list1->get_ListLevels()->idx_get(0)->get_StartAt());

        list2 = doc->get_Lists()->idx_get(1);
        TestUtil::VerifyListLevel(u"\u0000)", 18.0, NumberStyle::Arabic, list2->get_ListLevels()->idx_get(0));
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), list2->get_ListLevels()->idx_get(0)->get_Font()->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(10.0, list2->get_ListLevels()->idx_get(0)->get_Font()->get_Size());
        ASSERT_EQ(10, list2->get_ListLevels()->idx_get(0)->get_StartAt());
    }

    void CreateAndUseListStyle()
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
        auto doc = MakeObject<Document>();

        // Create a new list style
        // List formatting associated with this list style is default numbered
        SharedPtr<Style> listStyle = doc->get_Styles()->Add(StyleType::List, u"MyListStyle");

        // This list defines the formatting of the list style
        // Note that this list can not be used directly to apply formatting to paragraphs (see below)
        SharedPtr<List> list1 = listStyle->get_List();

        // Check some basic rules about the list that defines a list style
        ASSERT_TRUE(list1->get_IsListStyleDefinition());
        ASSERT_FALSE(list1->get_IsListStyleReference());
        ASSERT_TRUE(list1->get_IsMultiLevel());
        ASPOSE_ASSERT_EQ(listStyle, list1->get_Style());

        // Modify formatting of the list style to our liking
        for (auto level : System::IterateOver(list1->get_ListLevels()))
        {
            level->get_Font()->set_Name(u"Verdana");
            level->get_Font()->set_Color(System::Drawing::Color::get_Blue());
            level->get_Font()->set_Bold(true);
        }

        // Add some text to our document and use the list style
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Using list style first time:");

        // This creates a list based on the list style
        SharedPtr<List> list2 = doc->get_Lists()->Add(listStyle);

        // Check some basic rules about the list that references a list style
        ASSERT_FALSE(list2->get_IsListStyleDefinition());
        ASSERT_TRUE(list2->get_IsListStyleReference());
        ASPOSE_ASSERT_EQ(listStyle, list2->get_Style());

        // Apply the list that references the list style
        builder->get_ListFormat()->set_List(list2);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        builder->Writeln(u"Using list style second time:");

        // Create and apply another list based on the list style
        SharedPtr<List> list3 = doc->get_Lists()->Add(listStyle);
        builder->get_ListFormat()->set_List(list3);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        builder->get_Document()->Save(ArtifactsDir + u"Lists.CreateAndUseListStyle.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.CreateAndUseListStyle.docx");

        list1 = doc->get_Lists()->idx_get(0);

        TestUtil::VerifyListLevel(u"\u0000.", 18.0, NumberStyle::Arabic, list1->get_ListLevels()->idx_get(0));
        ASSERT_TRUE(list1->get_IsListStyleDefinition());
        ASSERT_FALSE(list1->get_IsListStyleReference());
        ASSERT_TRUE(list1->get_IsMultiLevel());
        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), list1->get_ListLevels()->idx_get(0)->get_Font()->get_Color().ToArgb());
        ASSERT_EQ(u"Verdana", list1->get_ListLevels()->idx_get(0)->get_Font()->get_Name());
        ASSERT_TRUE(list1->get_ListLevels()->idx_get(0)->get_Font()->get_Bold());

        list2 = doc->get_Lists()->idx_get(1);

        TestUtil::VerifyListLevel(u"\u0000.", 18.0, NumberStyle::Arabic, list2->get_ListLevels()->idx_get(0));
        ASSERT_FALSE(list2->get_IsListStyleDefinition());
        ASSERT_TRUE(list2->get_IsListStyleReference());
        ASSERT_TRUE(list2->get_IsMultiLevel());

        list3 = doc->get_Lists()->idx_get(2);

        TestUtil::VerifyListLevel(u"\u0000.", 18.0, NumberStyle::Arabic, list3->get_ListLevels()->idx_get(0));
        ASSERT_FALSE(list3->get_IsListStyleDefinition());
        ASSERT_TRUE(list3->get_IsListStyleReference());
        ASSERT_TRUE(list3->get_IsMultiLevel());
    }

    void DetectBulletedParagraphs()
    {
        //ExStart
        //ExFor:Paragraph.ListFormat
        //ExFor:ListFormat.IsListItem
        //ExFor:CompositeNode.GetText
        //ExFor:List.ListId
        //ExSummary:Shows how to output all paragraphs in a document that are bulleted or numbered.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

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

        SharedPtr<NodeCollection> paras = doc->GetChildNodes(NodeType::Paragraph, true);

        for (auto para : System::IterateOver(paras->LINQ_OfType<SharedPtr<Paragraph>>()->LINQ_Where([](SharedPtr<Paragraph> p) { return p->get_ListFormat()->get_IsListItem(); })))
        {
            std::cout << String::Format(u"This paragraph belongs to list ID# {0}, number style \"{1}\"",
                                                para->get_ListFormat()->get_List()->get_ListId(), para->get_ListFormat()->get_ListLevel()->get_NumberStyle())
                      << std::endl;
            std::cout << "\t\"" << para->GetText().Trim() << "\"" << std::endl;
        }
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        paras = doc->GetChildNodes(NodeType::Paragraph, true);

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> n)> _local_func_1 = [](SharedPtr<Node> n)
        {
            return (System::DynamicCast_noexcept<Paragraph>(n))->get_ListFormat()->get_IsListItem();
        };

        ASSERT_EQ(6, paras->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_1)));
    }

    void RemoveBulletsFromParagraphs()
    {
        //ExStart
        //ExFor:ListFormat.RemoveNumbers
        //ExSummary:Shows how to remove bullets and numbering from all paragraphs in the main text of a section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Numbered list item 1");
        builder->Writeln(u"Numbered list item 2");
        builder->Writeln(u"Numbered list item 3");
        builder->get_ListFormat()->RemoveNumbers();

        SharedPtr<NodeCollection> paras = doc->GetChildNodes(NodeType::Paragraph, true);

        ASSERT_EQ(3, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_IsListItem(); }));

        for (auto paragraph : System::IterateOver<Paragraph>(paras))
        {
            paragraph->get_ListFormat()->RemoveNumbers();
        }

        ASSERT_EQ(0, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_IsListItem(); }));
        //ExEnd
    }

    void ApplyExistingListToParagraphs()
    {
        //ExStart
        //ExFor:ListCollection.Item(Int32)
        //ExSummary:Shows how to apply list formatting of an existing list to a collection of paragraphs.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Paragraph 1");
        builder->Writeln(u"Paragraph 2");
        builder->Write(u"Paragraph 3");

        SharedPtr<NodeCollection> paras = doc->GetChildNodes(NodeType::Paragraph, true);

        ASSERT_EQ(0, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_IsListItem(); }));

        doc->get_Lists()->Add(ListTemplate::NumberDefault);
        SharedPtr<List> list = doc->get_Lists()->idx_get(0);

        for (auto paragraph : System::IterateOver(paras->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            paragraph->get_ListFormat()->set_List(list);
            paragraph->get_ListFormat()->set_ListLevelNumber(2);
        }

        ASSERT_EQ(3, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_IsListItem(); }));
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        paras = doc->GetChildNodes(NodeType::Paragraph, true);

        ASSERT_EQ(3, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_IsListItem(); }));

        ASSERT_EQ(3, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_ListLevelNumber() == 2; }));
    }

    void ApplyNewListToParagraphs()
    {
        //ExStart
        //ExFor:ListCollection.Add(ListTemplate)
        //ExSummary:Shows how to create a list by applying a new list format to a collection of paragraphs.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Paragraph 1");
        builder->Writeln(u"Paragraph 2");
        builder->Write(u"Paragraph 3");

        SharedPtr<NodeCollection> paras = doc->GetChildNodes(NodeType::Paragraph, true);

        ASSERT_EQ(0, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_IsListItem(); }));

        SharedPtr<List> list = doc->get_Lists()->Add(ListTemplate::NumberUppercaseLetterDot);

        for (auto paragraph : System::IterateOver(paras->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            paragraph->get_ListFormat()->set_List(list);
            paragraph->get_ListFormat()->set_ListLevelNumber(1);
        }

        ASSERT_EQ(3, paras->LINQ_Count([](SharedPtr<Node> n) { return System::DynamicCast<Paragraph>(n)->get_ListFormat()->get_IsListItem(); }));
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        paras = doc->GetChildNodes(NodeType::Paragraph, true);

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> n)> _local_func_10 = [](SharedPtr<Node> n)
        {
            return (System::DynamicCast_noexcept<Paragraph>(n))->get_ListFormat()->get_IsListItem();
        };

        ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_10)));

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Node> n)> _local_func_11 = [](SharedPtr<Node> n)
        {
            return (System::DynamicCast_noexcept<Paragraph>(n))->get_ListFormat()->get_ListLevelNumber() == 1;
        };

        ASSERT_EQ(3, paras->LINQ_Count(static_cast<System::Func<SharedPtr<Node>, bool>>(_local_func_11)));
    }

    //ExStart
    //ExFor:ListTemplate
    //ExSummary:Shows how to create a document that demonstrates all outline headings list templates.
    void OutlineHeadingTemplates()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<List> list = doc->get_Lists()->Add(ListTemplate::OutlineHeadingsArticleSection);
        AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Article Section\"");

        list = doc->get_Lists()->Add(ListTemplate::OutlineHeadingsLegal);
        AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Legal\"");

        builder->InsertBreak(BreakType::PageBreak);

        list = doc->get_Lists()->Add(ListTemplate::OutlineHeadingsNumbers);
        AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Numbers\"");

        list = doc->get_Lists()->Add(ListTemplate::OutlineHeadingsChapter);
        AddOutlineHeadingParagraphs(builder, list, u"Aspose.Words Outline - \"Chapters\"");

        doc->Save(ArtifactsDir + u"Lists.OutlineHeadingTemplates.docx");
        TestOutlineHeadingTemplates(MakeObject<Document>(ArtifactsDir + u"Lists.OutlineHeadingTemplates.docx"));
        //ExSkip
    }

protected:
    static void AddOutlineHeadingParagraphs(SharedPtr<DocumentBuilder> builder, SharedPtr<List> list, String title)
    {
        builder->get_ParagraphFormat()->ClearFormatting();
        builder->Writeln(title);

        for (int i = 0; i < 9; i++)
        {
            builder->get_ListFormat()->set_List(list);
            builder->get_ListFormat()->set_ListLevelNumber(i);

            String styleName = String(u"Heading ") + (i + 1);
            builder->get_ParagraphFormat()->set_StyleName(styleName);
            builder->Writeln(styleName);
        }

        builder->get_ListFormat()->RemoveNumbers();
    }
    //ExEnd

    void TestOutlineHeadingTemplates(SharedPtr<Document> doc)
    {
        SharedPtr<List> list = doc->get_Lists()->idx_get(0);
        // Article section list template

        TestUtil::VerifyListLevel(u"Article \u0000.", 0.0, NumberStyle::UppercaseRoman, list->get_ListLevels()->idx_get(0));
        TestUtil::VerifyListLevel(u"Section \u0000.\u0001", 0.0, NumberStyle::LeadingZero, list->get_ListLevels()->idx_get(1));
        TestUtil::VerifyListLevel(u"(\u0002)", 14.4, NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(2));
        TestUtil::VerifyListLevel(u"(\u0003)", 36.0, NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(3));
        TestUtil::VerifyListLevel(u"\u0004)", 28.8, NumberStyle::Arabic, list->get_ListLevels()->idx_get(4));
        TestUtil::VerifyListLevel(u"\u0005)", 36.0, NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(5));
        TestUtil::VerifyListLevel(u"\u0006)", 50.4, NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(6));
        TestUtil::VerifyListLevel(u"\a.", 50.4, NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(7));
        TestUtil::VerifyListLevel(u"\b.", 72.0, NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(8));

        list = doc->get_Lists()->idx_get(1);
        // Legal list template

        TestUtil::VerifyListLevel(u"\u0000", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(0));
        TestUtil::VerifyListLevel(u"\u0000.\u0001", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(1));
        TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(2));
        TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(3));
        TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(4));
        TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(5));
        TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(6));
        TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(7));
        TestUtil::VerifyListLevel(u"\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a.\b", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(8));

        list = doc->get_Lists()->idx_get(2);
        // Numbered list template

        TestUtil::VerifyListLevel(u"\u0000.", 0.0, NumberStyle::UppercaseRoman, list->get_ListLevels()->idx_get(0));
        TestUtil::VerifyListLevel(u"\u0001.", 36.0, NumberStyle::UppercaseLetter, list->get_ListLevels()->idx_get(1));
        TestUtil::VerifyListLevel(u"\u0002.", 72.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(2));
        TestUtil::VerifyListLevel(u"\u0003)", 108.0, NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(3));
        TestUtil::VerifyListLevel(u"(\u0004)", 144.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(4));
        TestUtil::VerifyListLevel(u"(\u0005)", 180.0, NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(5));
        TestUtil::VerifyListLevel(u"(\u0006)", 216.0, NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(6));
        TestUtil::VerifyListLevel(u"(\a)", 252.0, NumberStyle::LowercaseLetter, list->get_ListLevels()->idx_get(7));
        TestUtil::VerifyListLevel(u"(\b)", 288.0, NumberStyle::LowercaseRoman, list->get_ListLevels()->idx_get(8));

        list = doc->get_Lists()->idx_get(3);
        // Chapter list template

        TestUtil::VerifyListLevel(u"Chapter \u0000", 0.0, NumberStyle::Arabic, list->get_ListLevels()->idx_get(0));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(1));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(2));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(3));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(4));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(5));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(6));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(7));
        TestUtil::VerifyListLevel(u"", 0.0, NumberStyle::None, list->get_ListLevels()->idx_get(8));
    }

public:
    //ExStart
    //ExFor:ListCollection
    //ExFor:ListCollection.AddCopy(List)
    //ExFor:ListCollection.GetEnumerator
    //ExSummary:Shows how to enumerate through all lists defined in one document and creates a sample of those lists in another document.
    void PrintOutAllLists()
    {
        // Open a document that contains lists
        auto srcDoc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // This will be the sample document we product
        auto dstDoc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(dstDoc);

        for (auto srcList : System::IterateOver(srcDoc->get_Lists()))
        {
            // This copies the list formatting from the source into the destination document
            SharedPtr<List> dstList = dstDoc->get_Lists()->AddCopy(srcList);
            AddListSample(builder, dstList);
        }

        dstDoc->Save(ArtifactsDir + u"Lists.PrintOutAllLists.docx");
        TestPrintOutAllLists(srcDoc, MakeObject<Document>(ArtifactsDir + u"Lists.PrintOutAllLists.docx"));
        //ExSkip
    }

protected:
    static void AddListSample(SharedPtr<DocumentBuilder> builder, SharedPtr<List> list)
    {
        builder->Writeln(String(u"Sample formatting of list with ListId:") + list->get_ListId());
        builder->get_ListFormat()->set_List(list);
        for (int i = 0; i < list->get_ListLevels()->get_Count(); i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(String(u"Level ") + i);
        }

        builder->get_ListFormat()->RemoveNumbers();
        builder->Writeln();
    }
    //ExEnd

    void TestPrintOutAllLists(SharedPtr<Document> listSourceDoc, SharedPtr<Document> outDoc)
    {
        for (auto list : System::IterateOver(outDoc->get_Lists()))
        {
            for (int i = 0; i < list->get_ListLevels()->get_Count(); i++)
            {

                // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
                std::function<bool(SharedPtr<Aspose::Words::Lists::List> l)> _local_func_12 = [&list](SharedPtr<Aspose::Words::Lists::List> l)
                {
                    return l->get_ListId() == list->get_ListId();
                };

                SharedPtr<ListLevel> expectedListLevel = listSourceDoc->get_Lists()
                                                                     ->LINQ_First(static_cast<System::Func<SharedPtr<List>, bool>>(_local_func_12))
                                                                     ->get_ListLevels()
                                                                     ->idx_get(i);
                ASSERT_EQ(expectedListLevel->get_NumberFormat(), list->get_ListLevels()->idx_get(i)->get_NumberFormat());
                ASPOSE_ASSERT_EQ(expectedListLevel->get_NumberPosition(), list->get_ListLevels()->idx_get(i)->get_NumberPosition());
                ASSERT_EQ(expectedListLevel->get_NumberStyle(), list->get_ListLevels()->idx_get(i)->get_NumberStyle());
            }
        }
    }

public:
    void ListDocument()
    {
        //ExStart
        //ExFor:ListCollection.Document
        //ExFor:ListCollection.Count
        //ExFor:ListCollection.Item(Int32)
        //ExFor:ListCollection.GetListByListId
        //ExFor:List.Document
        //ExFor:List.ListId
        //ExSummary:Shows how to verify owner document properties of lists.
        auto doc = MakeObject<Document>();

        SharedPtr<ListCollection> lists = doc->get_Lists();

        ASPOSE_ASSERT_EQ(doc, lists->get_Document());

        SharedPtr<List> list = lists->Add(ListTemplate::BulletDefault);

        ASPOSE_ASSERT_EQ(doc, list->get_Document());

        std::cout << (String(u"Current list count: ") + lists->get_Count()) << std::endl;
        std::cout << (String(u"Is the first document list: ") + (System::ObjectExt::Equals(lists->idx_get(0), list))) << std::endl;
        std::cout << (String(u"ListId: ") + list->get_ListId()) << std::endl;
        std::cout << (String(u"List is the same by ListId: ") + (System::ObjectExt::Equals(lists->GetListByListId(1), list))) << std::endl;
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);
        lists = doc->get_Lists();

        ASPOSE_ASSERT_EQ(doc, lists->get_Document());
        ASSERT_EQ(1, lists->get_Count());
        ASSERT_EQ(1, lists->idx_get(0)->get_ListId());
        ASPOSE_ASSERT_EQ(lists->idx_get(0), lists->GetListByListId(1));
    }

    void CreateListRestartAfterHigher()
    {
        //ExStart
        //ExFor:ListLevel.NumberStyle
        //ExFor:ListLevel.NumberFormat
        //ExFor:ListLevel.IsLegal
        //ExFor:ListLevel.RestartAfterLevel
        //ExFor:ListLevel.LinkedStyle
        //ExFor:ListLevelCollection.GetEnumerator
        //ExSummary:Shows how to create a list with some advanced formatting.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<List> list = doc->get_Lists()->Add(ListTemplate::NumberDefault);

        // Level 1 labels will be "Appendix A", continuous and linked to the Heading 1 paragraph style
        list->get_ListLevels()->idx_get(0)->set_NumberFormat(u"Appendix \x0000");
        list->get_ListLevels()->idx_get(0)->set_NumberStyle(NumberStyle::UppercaseLetter);
        list->get_ListLevels()->idx_get(0)->set_LinkedStyle(doc->get_Styles()->idx_get(u"Heading 1"));

        // Level 2 labels will be "Section (1.01)" and restarting after Level 2 item appears
        list->get_ListLevels()->idx_get(1)->set_NumberFormat(u"Section (\x0000"
                                                             u".\x0001"
                                                             u")");
        list->get_ListLevels()->idx_get(1)->set_NumberStyle(NumberStyle::LeadingZero);

        // Note that the higher level uses UppercaseLetter numbering
        // We can set this property to use Arabic numbers for the higher list levels
        list->get_ListLevels()->idx_get(1)->set_IsLegal(true);
        list->get_ListLevels()->idx_get(1)->set_RestartAfterLevel(0);

        // Level 3 labels will be "-I-" and restarting after Level 2 item appears
        list->get_ListLevels()->idx_get(2)->set_NumberFormat(u"-\x0002"
                                                             u"-");
        list->get_ListLevels()->idx_get(2)->set_NumberStyle(NumberStyle::UppercaseRoman);
        list->get_ListLevels()->idx_get(2)->set_RestartAfterLevel(1);

        // Make labels of all list levels bold
        for (auto level : System::IterateOver(list->get_ListLevels()))
        {
            level->get_Font()->set_Bold(true);
        }

        // Apply list formatting to the current paragraph
        builder->get_ListFormat()->set_List(list);

        // Exercise the 3 levels we created two times
        for (int n = 0; n < 2; n++)
        {
            for (int i = 0; i < 3; i++)
            {
                builder->get_ListFormat()->set_ListLevelNumber(i);
                builder->Writeln(String(u"Level ") + i);
            }
        }

        builder->get_ListFormat()->RemoveNumbers();

        doc->Save(ArtifactsDir + u"Lists.CreateListRestartAfterHigher.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.CreateListRestartAfterHigher.docx");

        SharedPtr<ListLevel> listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0);

        TestUtil::VerifyListLevel(u"Appendix \u0000", 18.0, NumberStyle::UppercaseLetter, listLevel);
        ASSERT_FALSE(listLevel->get_IsLegal());
        ASSERT_EQ(-1, listLevel->get_RestartAfterLevel());
        ASSERT_EQ(u"Heading 1", listLevel->get_LinkedStyle()->get_Name());

        listLevel = doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(1);

        TestUtil::VerifyListLevel(u"Section (\u0000.\u0001)", 54.0, NumberStyle::LeadingZero, listLevel);
        ASSERT_TRUE(listLevel->get_IsLegal());
        ASSERT_EQ(0, listLevel->get_RestartAfterLevel());
        ASSERT_TRUE(listLevel->get_LinkedStyle() == nullptr);
    }

    void GetListLabels()
    {
        //ExStart
        //ExFor:Document.UpdateListLabels()
        //ExFor:Node.ToString(SaveFormat)
        //ExFor:ListLabel
        //ExFor:Paragraph.ListLabel
        //ExFor:ListLabel.LabelValue
        //ExFor:ListLabel.LabelString
        //ExSummary:Shows how to extract the label of each paragraph in a list as a value or a String.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        doc->UpdateListLabels();

        SharedPtr<NodeCollection> paras = doc->GetChildNodes(NodeType::Paragraph, true);

        // Find if we have the paragraph list. In our document our list uses plain Arabic numbers,
        // which start at three and ends at six

        for (auto paragraph : System::IterateOver(paras->LINQ_OfType<SharedPtr<Paragraph>>()->LINQ_Where([](SharedPtr<Paragraph> p) { return p->get_ListFormat()->get_IsListItem(); })))
        {
            std::cout << "List item paragraph #" << paras->IndexOf(paragraph) << std::endl;

            // This is the text we get when actually getting when we output this node to text format
            // The list labels are not included in this text output. Trim any paragraph formatting characters
            String paragraphText = paragraph->ToString(SaveFormat::Text).Trim();
            std::cout << "\tExported Text: " << paragraphText << std::endl;

            SharedPtr<ListLabel> label = paragraph->get_ListLabel();

            // This gets the position of the paragraph in current level of the list. If we have a list with multiple levels,
            // this will tell us what position it is on that level
            std::cout << "\tNumerical Id: " << label->get_LabelValue() << std::endl;

            // Combine them together to include the list label with the text in the output
            std::cout << "\tList label combined with text: " << label->get_LabelString() << " " << paragraphText << std::endl;
        }
        //ExEnd

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<bool(SharedPtr<Paragraph> p)> _local_func_14 = [](SharedPtr<Paragraph> p)
        {
            return p->get_ListFormat()->get_IsListItem();
        };

        ASSERT_EQ(
            10, paras->LINQ_OfType<SharedPtr<Paragraph>>()->LINQ_Count(static_cast<System::Func<SharedPtr<Paragraph>, bool>>(_local_func_14)));
    }

    void CreatePictureBullet()
    {
        //ExStart
        //ExFor:ListLevel.CreatePictureBullet
        //ExFor:ListLevel.DeletePictureBullet
        //ExSummary:Shows how to creating and deleting picture bullet with custom image.
        auto doc = MakeObject<Document>();

        // Create a list with template
        SharedPtr<List> list = doc->get_Lists()->Add(ListTemplate::BulletCircle);

        // Create picture bullet for the current list level
        list->get_ListLevels()->idx_get(0)->CreatePictureBullet();

        // Set your own picture bullet image through the ImageData
        list->get_ListLevels()->idx_get(0)->get_ImageData()->SetImage(ImageDir + u"Logo icon.ico");

        ASSERT_TRUE(list->get_ListLevels()->idx_get(0)->get_ImageData()->get_HasImage());

        // Create a list, configure its bullets to use our image and add two list items
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_ListFormat()->set_List(list);
        builder->Writeln(u"Hello world!");
        builder->Write(u"Hello again!");

        doc->Save(ArtifactsDir + u"Lists.CreatePictureBullet.docx");

        // Delete picture bullet
        list->get_ListLevels()->idx_get(0)->DeletePictureBullet();

        ASSERT_TRUE(list->get_ListLevels()->idx_get(0)->get_ImageData() == nullptr);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Lists.CreatePictureBullet.docx");

        ASSERT_TRUE(doc->get_Lists()->idx_get(0)->get_ListLevels()->idx_get(0)->get_ImageData()->get_HasImage());
    }
};

} // namespace ApiExamples
