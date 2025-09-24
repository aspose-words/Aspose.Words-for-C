// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExStyles.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/linq/enumerable.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/collections/ienumerator.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/TabStopCollection.h>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
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
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

#include "DocumentHelper.h"


using namespace Aspose::Words::Lists;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1958518354u, ::Aspose::Words::ApiExamples::ExStyles, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExStyles : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExStyles> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExStyles>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExStyles> ExStyles::s_instance;

} // namespace gtest_test

void ExStyles::Styles()
{
    //ExStart
    //ExFor:DocumentBase.Styles
    //ExFor:Style.Document
    //ExFor:Style.Name
    //ExFor:Style.IsHeading
    //ExFor:Style.IsQuickStyle
    //ExFor:Style.NextParagraphStyleName
    //ExFor:Style.Styles
    //ExFor:Style.Type
    //ExFor:StyleCollection.Document
    //ExFor:StyleCollection.GetEnumerator
    //ExSummary:Shows how to access a document's style collection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_EQ(4, doc->get_Styles()->get_Count());
    
    // Enumerate and list all the styles that a document created using Aspose.Words contains by default.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Style>>> stylesEnum = doc->get_Styles()->GetEnumerator();
        while (stylesEnum->MoveNext())
        {
            System::SharedPtr<Aspose::Words::Style> curStyle = stylesEnum->get_Current();
            std::cout << System::String::Format(u"Style name:\t\"{0}\", of type \"{1}\"", curStyle->get_Name(), curStyle->get_Type()) << std::endl;
            std::cout << System::String::Format(u"\tSubsequent style:\t{0}", curStyle->get_NextParagraphStyleName()) << std::endl;
            std::cout << System::String::Format(u"\tIs heading:\t\t\t{0}", curStyle->get_IsHeading()) << std::endl;
            std::cout << System::String::Format(u"\tIs QuickStyle:\t\t{0}", curStyle->get_IsQuickStyle()) << std::endl;
            
            ASPOSE_ASSERT_EQ(doc, curStyle->get_Document());
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStyles, Styles)
{
    s_instance->Styles();
}

} // namespace gtest_test

void ExStyles::CreateStyle()
{
    //ExStart
    //ExFor:Style.Font
    //ExFor:Style
    //ExFor:Style.Remove
    //ExFor:Style.AutomaticallyUpdate
    //ExSummary:Shows how to create and apply a custom style.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Style> style = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyStyle");
    style->get_Font()->set_Name(u"Times New Roman");
    style->get_Font()->set_Size(16);
    style->get_Font()->set_Color(System::Drawing::Color::get_Navy());
    // Automatically redefine style.
    style->set_AutomaticallyUpdate(true);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Apply one of the styles from the document to the paragraph that the document builder is creating.
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"MyStyle"));
    builder->Writeln(u"Hello world!");
    
    System::SharedPtr<Aspose::Words::Style> firstParagraphStyle = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Style();
    
    ASPOSE_ASSERT_EQ(style, firstParagraphStyle);
    
    // Remove our custom style from the document's styles collection.
    doc->get_Styles()->idx_get(u"MyStyle")->Remove();
    
    firstParagraphStyle = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Style();
    
    // Any text that used a removed style reverts to the default formatting.
    ASSERT_FALSE(doc->get_Styles()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Style>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Style> s)>>([](System::SharedPtr<Aspose::Words::Style> s) -> bool
    {
        return s->get_Name() == u"MyStyle";
    }))));
    ASSERT_EQ(u"Times New Roman", firstParagraphStyle->get_Font()->get_Name());
    ASPOSE_ASSERT_EQ(12.0, firstParagraphStyle->get_Font()->get_Size());
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), firstParagraphStyle->get_Font()->get_Color().ToArgb());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStyles, CreateStyle)
{
    s_instance->CreateStyle();
}

} // namespace gtest_test

void ExStyles::StyleCollection()
{
    //ExStart
    //ExFor:StyleCollection.Add(StyleType,String)
    //ExFor:StyleCollection.Count
    //ExFor:StyleCollection.DefaultFont
    //ExFor:StyleCollection.DefaultParagraphFormat
    //ExFor:StyleCollection.Item(StyleIdentifier)
    //ExFor:StyleCollection.Item(Int32)
    //ExSummary:Shows how to add a Style to a document's styles collection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::StyleCollection> styles = doc->get_Styles();
    // Set default parameters for new styles that we may later add to this collection.
    styles->get_DefaultFont()->set_Name(u"Courier New");
    // If we add a style of the "StyleType.Paragraph", the collection will apply the values of
    // its "DefaultParagraphFormat" property to the style's "ParagraphFormat" property.
    styles->get_DefaultParagraphFormat()->set_FirstLineIndent(15.0);
    // Add a style, and then verify that it has the default settings.
    styles->Add(Aspose::Words::StyleType::Paragraph, u"MyStyle");
    
    ASSERT_EQ(u"Courier New", styles->idx_get(4)->get_Font()->get_Name());
    ASPOSE_ASSERT_EQ(15.0, styles->idx_get(u"MyStyle")->get_ParagraphFormat()->get_FirstLineIndent());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStyles, StyleCollection)
{
    s_instance->StyleCollection();
}

} // namespace gtest_test

void ExStyles::RemoveStylesFromStyleGallery()
{
    //ExStart
    //ExFor:StyleCollection.ClearQuickStyleGallery
    //ExSummary:Shows how to remove styles from Style Gallery panel.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    // Note that remove styles work only with DOCX format for now.
    doc->get_Styles()->ClearQuickStyleGallery();
    
    doc->Save(get_ArtifactsDir() + u"Styles.RemoveStylesFromStyleGallery.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStyles, RemoveStylesFromStyleGallery)
{
    s_instance->RemoveStylesFromStyleGallery();
}

} // namespace gtest_test

void ExStyles::ChangeTocsTabStops()
{
    //ExStart
    //ExFor:TabStop
    //ExFor:ParagraphFormat.TabStops
    //ExFor:Style.StyleIdentifier
    //ExFor:TabStopCollection.RemoveByPosition
    //ExFor:TabStop.Alignment
    //ExFor:TabStop.Position
    //ExFor:TabStop.Leader
    //ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Table of contents.docx");
    
    // Iterate through all paragraphs with TOC result-based styles; this is any style between TOC and TOC9.
    for (auto&& para : System::IterateOver<Aspose::Words::Paragraph>(doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)))
    {
        if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= Aspose::Words::StyleIdentifier::Toc1 && para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= Aspose::Words::StyleIdentifier::Toc9)
        {
            // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
            System::SharedPtr<Aspose::Words::TabStop> tab = para->get_ParagraphFormat()->get_TabStops()->idx_get(0);
            
            // Replace the first default tab, stop with a custom tab stop.
            para->get_ParagraphFormat()->get_TabStops()->RemoveByPosition(tab->get_Position());
            para->get_ParagraphFormat()->get_TabStops()->Add(tab->get_Position() - 50, tab->get_Alignment(), tab->get_Leader());
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"Styles.ChangeTocsTabStops.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Styles.ChangeTocsTabStops.docx");
    
    for (auto&& para : System::IterateOver<Aspose::Words::Paragraph>(doc->GetChildNodes(Aspose::Words::NodeType::Paragraph, true)))
    {
        if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= Aspose::Words::StyleIdentifier::Toc1 && para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= Aspose::Words::StyleIdentifier::Toc9)
        {
            System::SharedPtr<Aspose::Words::TabStop> tabStop = para->GetEffectiveTabStops()->idx_get(0);
            ASPOSE_ASSERT_EQ(400.8, tabStop->get_Position());
            ASSERT_EQ(Aspose::Words::TabAlignment::Right, tabStop->get_Alignment());
            ASSERT_EQ(Aspose::Words::TabLeader::Dots, tabStop->get_Leader());
        }
    }
}

namespace gtest_test
{

TEST_F(ExStyles, ChangeTocsTabStops)
{
    s_instance->ChangeTocsTabStops();
}

} // namespace gtest_test

void ExStyles::CopyStyleSameDocument()
{
    //ExStart
    //ExFor:StyleCollection.AddCopy(Style)
    //ExFor:Style.Name
    //ExSummary:Shows how to clone a document's style.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // The AddCopy method creates a copy of the specified style and
    // automatically generates a new name for the style, such as "Heading 1_0".
    System::SharedPtr<Aspose::Words::Style> newStyle = doc->get_Styles()->AddCopy(doc->get_Styles()->idx_get(u"Heading 1"));
    
    // Use the style's "Name" property to change the style's identifying name.
    newStyle->set_Name(u"My Heading 1");
    
    // Our document now has two identical looking styles with different names.
    // Changing settings of one of the styles do not affect the other.
    newStyle->get_Font()->set_Color(System::Drawing::Color::get_Red());
    
    ASSERT_EQ(u"My Heading 1", newStyle->get_Name());
    ASSERT_EQ(u"Heading 1", doc->get_Styles()->idx_get(u"Heading 1")->get_Name());
    
    ASSERT_EQ(doc->get_Styles()->idx_get(u"Heading 1")->get_Type(), newStyle->get_Type());
    ASSERT_EQ(doc->get_Styles()->idx_get(u"Heading 1")->get_Font()->get_Name(), newStyle->get_Font()->get_Name());
    ASPOSE_ASSERT_EQ(doc->get_Styles()->idx_get(u"Heading 1")->get_Font()->get_Size(), newStyle->get_Font()->get_Size());
    ASPOSE_ASSERT_NE(doc->get_Styles()->idx_get(u"Heading 1")->get_Font()->get_Color(), newStyle->get_Font()->get_Color());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStyles, CopyStyleSameDocument)
{
    s_instance->CopyStyleSameDocument();
}

} // namespace gtest_test

void ExStyles::CopyStyleDifferentDocument()
{
    //ExStart
    //ExFor:StyleCollection.AddCopy(Style)
    //ExSummary:Shows how to import a style from one document into a different document.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a custom style for the source document.
    System::SharedPtr<Aspose::Words::Style> srcStyle = srcDoc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyStyle");
    srcStyle->get_Font()->set_Color(System::Drawing::Color::get_Red());
    
    // Import the source document's custom style into the destination document.
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Style> newStyle = dstDoc->get_Styles()->AddCopy(srcStyle);
    
    // The imported style has an appearance identical to its source style.
    ASSERT_EQ(u"MyStyle", newStyle->get_Name());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), newStyle->get_Font()->get_Color().ToArgb());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStyles, CopyStyleDifferentDocument)
{
    s_instance->CopyStyleDifferentDocument();
}

} // namespace gtest_test

void ExStyles::DefaultStyles()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    doc->get_Styles()->get_DefaultFont()->set_Name(u"PMingLiU");
    doc->get_Styles()->get_DefaultFont()->set_Bold(true);
    
    doc->get_Styles()->get_DefaultParagraphFormat()->set_SpaceAfter(20);
    doc->get_Styles()->get_DefaultParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Right);
    
    doc = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(doc);
    
    ASSERT_TRUE(doc->get_Styles()->get_DefaultFont()->get_Bold());
    ASSERT_EQ(u"PMingLiU", doc->get_Styles()->get_DefaultFont()->get_Name());
    ASPOSE_ASSERT_EQ(20, doc->get_Styles()->get_DefaultParagraphFormat()->get_SpaceAfter());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, doc->get_Styles()->get_DefaultParagraphFormat()->get_Alignment());
}

namespace gtest_test
{

TEST_F(ExStyles, DefaultStyles)
{
    s_instance->DefaultStyles();
}

} // namespace gtest_test

void ExStyles::ParagraphStyleBulletedList()
{
    //ExStart
    //ExFor:StyleCollection
    //ExFor:DocumentBase.Styles
    //ExFor:Style
    //ExFor:Font
    //ExFor:Style.Font
    //ExFor:Style.ParagraphFormat
    //ExFor:Style.ListFormat
    //ExFor:ParagraphFormat.Style
    //ExSummary:Shows how to create and use a paragraph style with list formatting.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a custom paragraph style.
    System::SharedPtr<Aspose::Words::Style> style = doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"MyStyle1");
    style->get_Font()->set_Size(24);
    style->get_Font()->set_Name(u"Verdana");
    style->get_ParagraphFormat()->set_SpaceAfter(12);
    
    // Create a list and make sure the paragraphs that use this style will use this list.
    style->get_ListFormat()->set_List(doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::BulletDefault));
    style->get_ListFormat()->set_ListLevelNumber(0);
    
    // Apply the paragraph style to the document builder's current paragraph, and then add some text.
    builder->get_ParagraphFormat()->set_Style(style);
    builder->Writeln(u"Hello World: MyStyle1, bulleted list.");
    
    // Change the document builder's style to one that has no list formatting and write another paragraph.
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
    builder->Writeln(u"Hello World: Normal.");
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"Styles.ParagraphStyleBulletedList.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Styles.ParagraphStyleBulletedList.docx");
    
    style = doc->get_Styles()->idx_get(u"MyStyle1");
    
    ASSERT_EQ(u"MyStyle1", style->get_Name());
    ASPOSE_ASSERT_EQ(24, style->get_Font()->get_Size());
    ASSERT_EQ(u"Verdana", style->get_Font()->get_Name());
    ASPOSE_ASSERT_EQ(12.0, style->get_ParagraphFormat()->get_SpaceAfter());
}

namespace gtest_test
{

TEST_F(ExStyles, ParagraphStyleBulletedList)
{
    s_instance->ParagraphStyleBulletedList();
}

} // namespace gtest_test

void ExStyles::StyleAliases()
{
    //ExStart
    //ExFor:Style.Aliases
    //ExFor:Style.BaseStyleName
    //ExFor:Style.Equals(Style)
    //ExFor:Style.LinkedStyleName
    //ExSummary:Shows how to use style aliases.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Style with alias.docx");
    
    // This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
    // If a style's name has multiple values separated by commas, each clause is a separate alias.
    System::SharedPtr<Aspose::Words::Style> style = doc->get_Styles()->idx_get(u"MyStyle");
    ASPOSE_ASSERT_EQ(System::MakeArray<System::String>({u"MyStyle Alias 1", u"MyStyle Alias 2"}), style->get_Aliases());
    ASSERT_EQ(u"Title", style->get_BaseStyleName());
    ASSERT_EQ(u"MyStyle Char", style->get_LinkedStyleName());
    
    // We can reference a style using its alias, as well as its name.
    ASPOSE_ASSERT_EQ(doc->get_Styles()->idx_get(u"MyStyle Alias 1"), doc->get_Styles()->idx_get(u"MyStyle Alias 2"));
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"MyStyle Alias 1"));
    builder->Writeln(u"Hello world!");
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"MyStyle Alias 2"));
    builder->Write(u"Hello again!");
    
    ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Style(), doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_Style());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExStyles, StyleAliases)
{
    s_instance->StyleAliases();
}

} // namespace gtest_test

void ExStyles::LockStyle()
{
    //ExStart:LockStyle
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:Style.Locked
    //ExSummary:Shows how to lock style.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Style> styleHeading1 = doc->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Heading1);
    if (!styleHeading1->get_Locked())
    {
        styleHeading1->set_Locked(true);
    }
    
    doc->Save(get_ArtifactsDir() + u"Styles.LockStyle.docx");
    //ExEnd:LockStyle
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Styles.LockStyle.docx");
    ASSERT_TRUE(doc->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Heading1)->get_Locked());
}

namespace gtest_test
{

TEST_F(ExStyles, LockStyle)
{
    s_instance->LockStyle();
}

} // namespace gtest_test

void ExStyles::StylePriority()
{
    //ExStart:StylePriority
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:Style.Priority
    //ExFor:Style.UnhideWhenUsed
    //ExFor:Style.SemiHidden
    //ExSummary:Shows how to prioritize and hide a style.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Style> styleTitle = doc->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Subtitle);
    
    if (styleTitle->get_Priority() == 9)
    {
        styleTitle->set_Priority(10);
    }
    
    if (!styleTitle->get_UnhideWhenUsed())
    {
        styleTitle->set_UnhideWhenUsed(true);
    }
    
    if (styleTitle->get_SemiHidden())
    {
        styleTitle->set_SemiHidden(true);
    }
    
    doc->Save(get_ArtifactsDir() + u"Styles.StylePriority.docx");
    //ExEnd:StylePriority
}

namespace gtest_test
{

TEST_F(ExStyles, StylePriority)
{
    s_instance->StylePriority();
}

} // namespace gtest_test

void ExStyles::LinkedStyleName()
{
    //ExStart:LinkedStyleName
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:Style.LinkedStyleName
    //ExSummary:Shows how to link styles among themselves.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Style> styleHeading1 = doc->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Heading1);
    
    System::SharedPtr<Aspose::Words::Style> styleHeading1Char = doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"Heading 1 Char");
    styleHeading1Char->get_Font()->set_Name(u"Verdana");
    styleHeading1Char->get_Font()->set_Bold(true);
    styleHeading1Char->get_Font()->get_Border()->set_LineStyle(Aspose::Words::LineStyle::Dot);
    styleHeading1Char->get_Font()->get_Border()->set_LineWidth(15);
    
    styleHeading1->set_LinkedStyleName(u"Heading 1 Char");
    
    ASSERT_EQ(u"Heading 1 Char", styleHeading1->get_LinkedStyleName());
    ASSERT_EQ(u"Heading 1", styleHeading1Char->get_LinkedStyleName());
    //ExEnd:LinkedStyleName
}

namespace gtest_test
{

TEST_F(ExStyles, LinkedStyleName)
{
    s_instance->LinkedStyleName();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
