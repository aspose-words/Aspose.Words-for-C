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
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <Aspose.Words.Cpp/TabAlignment.h>
#include <Aspose.Words.Cpp/TabLeader.h>
#include <Aspose.Words.Cpp/TabStop.h>
#include <Aspose.Words.Cpp/TabStopCollection.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/collections/ienumerator.h>
#include <system/convert.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;

namespace ApiExamples {

class ExStyles : public ApiExampleBase
{
public:
    void Styles()
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
        auto doc = MakeObject<Document>();

        ASSERT_EQ(4, doc->get_Styles()->get_Count());

        // Enumerate and list all the styles that a document created using Aspose.Words contains by default.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Style>>> stylesEnum = doc->get_Styles()->GetEnumerator();
            while (stylesEnum->MoveNext())
            {
                SharedPtr<Style> curStyle = stylesEnum->get_Current();
                std::cout << String::Format(u"Style name:\t\"{0}\", of type \"{1}\"", curStyle->get_Name(), curStyle->get_Type()) << std::endl;
                std::cout << "\tSubsequent style:\t" << curStyle->get_NextParagraphStyleName() << std::endl;
                std::cout << String::Format(u"\tIs heading:\t\t\t{0}", curStyle->get_IsHeading()) << std::endl;
                std::cout << String::Format(u"\tIs QuickStyle:\t\t{0}", curStyle->get_IsQuickStyle()) << std::endl;

                ASPOSE_ASSERT_EQ(doc, curStyle->get_Document());
            }
        }
        //ExEnd
    }

    void CreateStyle()
    {
        //ExStart
        //ExFor:Style.Font
        //ExFor:Style
        //ExFor:Style.Remove
        //ExSummary:Shows how to create and apply a custom style.
        auto doc = MakeObject<Document>();

        SharedPtr<Style> style = doc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle");
        style->get_Font()->set_Name(u"Times New Roman");
        style->get_Font()->set_Size(16);
        style->get_Font()->set_Color(System::Drawing::Color::get_Navy());

        auto builder = MakeObject<DocumentBuilder>(doc);

        // Apply one of the styles from the document to the paragraph that the document builder is creating.
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"MyStyle"));
        builder->Writeln(u"Hello world!");

        SharedPtr<Style> firstParagraphStyle = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Style();

        ASPOSE_ASSERT_EQ(style, firstParagraphStyle);

        // Remove our custom style from the document's styles collection.
        doc->get_Styles()->idx_get(u"MyStyle")->Remove();

        firstParagraphStyle = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Style();

        // Any text that used a removed style reverts to the default formatting.
        ASSERT_FALSE(doc->get_Styles()->LINQ_Any([](SharedPtr<Style> s) { return s->get_Name() == u"MyStyle"; }));
        ASSERT_EQ(u"Times New Roman", firstParagraphStyle->get_Font()->get_Name());
        ASPOSE_ASSERT_EQ(12.0, firstParagraphStyle->get_Font()->get_Size());
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), firstParagraphStyle->get_Font()->get_Color().ToArgb());
        //ExEnd
    }

    void StyleCollection_()
    {
        //ExStart
        //ExFor:StyleCollection.Add(StyleType,String)
        //ExFor:StyleCollection.Count
        //ExFor:StyleCollection.DefaultFont
        //ExFor:StyleCollection.DefaultParagraphFormat
        //ExFor:StyleCollection.Item(StyleIdentifier)
        //ExFor:StyleCollection.Item(Int32)
        //ExSummary:Shows how to add a Style to a document's styles collection.
        auto doc = MakeObject<Document>();
        SharedPtr<StyleCollection> styles = doc->get_Styles();

        // Set default parameters for new styles that we may later add to this collection.
        styles->get_DefaultFont()->set_Name(u"Courier New");

        // If we add a style of the "StyleType.Paragraph", the collection will apply the values of
        // its "DefaultParagraphFormat" property to the style's "ParagraphFormat" property.
        styles->get_DefaultParagraphFormat()->set_FirstLineIndent(15.0);

        // Add a style, and then verify that it has the default settings.
        styles->Add(StyleType::Paragraph, u"MyStyle");

        ASSERT_EQ(u"Courier New", styles->idx_get(4)->get_Font()->get_Name());
        ASPOSE_ASSERT_EQ(15.0, styles->idx_get(u"MyStyle")->get_ParagraphFormat()->get_FirstLineIndent());
        //ExEnd
    }

    void RemoveStylesFromStyleGallery()
    {
        //ExStart
        //ExFor:StyleCollection.ClearQuickStyleGallery
        //ExSummary:Shows how to remove styles from Style Gallery panel.
        auto doc = MakeObject<Document>();

        // Note that remove styles work only with DOCX format for now.
        doc->get_Styles()->ClearQuickStyleGallery();

        doc->Save(ArtifactsDir + u"Styles.RemoveStylesFromStyleGallery.docx");
        //ExEnd
    }

    void ChangeTocsTabStops()
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
        auto doc = MakeObject<Document>(MyDir + u"Table of contents.docx");

        // Iterate through all paragraphs with TOC result-based styles; this is any style between TOC and TOC9.
        for (auto para : System::IterateOver(doc->GetChildNodes(NodeType::Paragraph, true)->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= StyleIdentifier::Toc1 &&
                para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= StyleIdentifier::Toc9)
            {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                SharedPtr<TabStop> tab = para->get_ParagraphFormat()->get_TabStops()->idx_get(0);

                // Replace the first default tab, stop with a custom tab stop.
                para->get_ParagraphFormat()->get_TabStops()->RemoveByPosition(tab->get_Position());
                para->get_ParagraphFormat()->get_TabStops()->Add(tab->get_Position() - 50, tab->get_Alignment(), tab->get_Leader());
            }
        }

        doc->Save(ArtifactsDir + u"Styles.ChangeTocsTabStops.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Styles.ChangeTocsTabStops.docx");

        for (auto para : System::IterateOver(doc->GetChildNodes(NodeType::Paragraph, true)->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= StyleIdentifier::Toc1 &&
                para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= StyleIdentifier::Toc9)
            {
                SharedPtr<TabStop> tabStop = para->GetEffectiveTabStops()->idx_get(0);
                ASPOSE_ASSERT_EQ(400.8, tabStop->get_Position());
                ASSERT_EQ(TabAlignment::Right, tabStop->get_Alignment());
                ASSERT_EQ(TabLeader::Dots, tabStop->get_Leader());
            }
        }
    }

    void CopyStyleSameDocument()
    {
        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExFor:Style.Name
        //ExSummary:Shows how to clone a document's style.
        auto doc = MakeObject<Document>();

        // The AddCopy method creates a copy of the specified style and
        // automatically generates a new name for the style, such as "Heading 1_0".
        SharedPtr<Style> newStyle = doc->get_Styles()->AddCopy(doc->get_Styles()->idx_get(u"Heading 1"));

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

    void CopyStyleDifferentDocument()
    {
        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExSummary:Shows how to import a style from one document into a different document.
        auto srcDoc = MakeObject<Document>();

        // Create a custom style for the source document.
        SharedPtr<Style> srcStyle = srcDoc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle");
        srcStyle->get_Font()->set_Color(System::Drawing::Color::get_Red());

        // Import the source document's custom style into the destination document.
        auto dstDoc = MakeObject<Document>();
        SharedPtr<Style> newStyle = dstDoc->get_Styles()->AddCopy(srcStyle);

        // The imported style has an appearance identical to its source style.
        ASSERT_EQ(u"MyStyle", newStyle->get_Name());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), newStyle->get_Font()->get_Color().ToArgb());
        //ExEnd
    }

    void DefaultStyles()
    {
        auto doc = MakeObject<Document>();

        doc->get_Styles()->get_DefaultFont()->set_Name(u"PMingLiU");
        doc->get_Styles()->get_DefaultFont()->set_Bold(true);

        doc->get_Styles()->get_DefaultParagraphFormat()->set_SpaceAfter(20);
        doc->get_Styles()->get_DefaultParagraphFormat()->set_Alignment(ParagraphAlignment::Right);

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_TRUE(doc->get_Styles()->get_DefaultFont()->get_Bold());
        ASSERT_EQ(u"PMingLiU", doc->get_Styles()->get_DefaultFont()->get_Name());
        ASPOSE_ASSERT_EQ(20, doc->get_Styles()->get_DefaultParagraphFormat()->get_SpaceAfter());
        ASSERT_EQ(ParagraphAlignment::Right, doc->get_Styles()->get_DefaultParagraphFormat()->get_Alignment());
    }

    void ParagraphStyleBulletedList()
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
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a custom paragraph style.
        SharedPtr<Style> style = doc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle1");
        style->get_Font()->set_Size(24);
        style->get_Font()->set_Name(u"Verdana");
        style->get_ParagraphFormat()->set_SpaceAfter(12);

        // Create a list and make sure the paragraphs that use this style will use this list.
        style->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::BulletDefault));
        style->get_ListFormat()->set_ListLevelNumber(0);

        // Apply the paragraph style to the document builder's current paragraph, and then add some text.
        builder->get_ParagraphFormat()->set_Style(style);
        builder->Writeln(u"Hello World: MyStyle1, bulleted list.");

        // Change the document builder's style to one that has no list formatting and write another paragraph.
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
        builder->Writeln(u"Hello World: Normal.");

        builder->get_Document()->Save(ArtifactsDir + u"Styles.ParagraphStyleBulletedList.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Styles.ParagraphStyleBulletedList.docx");

        style = doc->get_Styles()->idx_get(u"MyStyle1");

        ASSERT_EQ(u"MyStyle1", style->get_Name());
        ASPOSE_ASSERT_EQ(24, style->get_Font()->get_Size());
        ASSERT_EQ(u"Verdana", style->get_Font()->get_Name());
        ASPOSE_ASSERT_EQ(12.0, style->get_ParagraphFormat()->get_SpaceAfter());
    }

    void StyleAliases()
    {
        //ExStart
        //ExFor:Style.Aliases
        //ExFor:Style.BaseStyleName
        //ExFor:Style.Equals(Aspose.Words.Style)
        //ExFor:Style.LinkedStyleName
        //ExSummary:Shows how to use style aliases.
        auto doc = MakeObject<Document>(MyDir + u"Style with alias.docx");

        // This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
        // If a style's name has multiple values separated by commas, each clause is a separate alias.
        SharedPtr<Style> style = doc->get_Styles()->idx_get(u"MyStyle");
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"MyStyle Alias 1", u"MyStyle Alias 2"}), style->get_Aliases());
        ASSERT_EQ(u"Title", style->get_BaseStyleName());
        ASSERT_EQ(u"MyStyle Char", style->get_LinkedStyleName());

        // We can reference a style using its alias, as well as its name.
        ASPOSE_ASSERT_EQ(doc->get_Styles()->idx_get(u"MyStyle Alias 1"), doc->get_Styles()->idx_get(u"MyStyle Alias 2"));

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->MoveToDocumentEnd();
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"MyStyle Alias 1"));
        builder->Writeln(u"Hello world!");
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"MyStyle Alias 2"));
        builder->Write(u"Hello again!");

        ASPOSE_ASSERT_EQ(doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(0)->get_ParagraphFormat()->get_Style(),
                         doc->get_FirstSection()->get_Body()->get_Paragraphs()->idx_get(1)->get_ParagraphFormat()->get_Style());
        //ExEnd
    }
};

} // namespace ApiExamples
