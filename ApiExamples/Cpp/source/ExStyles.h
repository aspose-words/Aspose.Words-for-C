#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabStopCollection.h>
#include <drawing/color.h>
#include <system/array.h>
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

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

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

        // A blank document comes with 4 styles by default
        ASSERT_EQ(4, doc->get_Styles()->get_Count());

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
        //ExSummary:Shows how to create and apply a style.
        auto doc = MakeObject<Document>();

        // Add a custom style and change its appearance
        SharedPtr<Style> style = doc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle");
        style->get_Font()->set_Name(u"Times New Roman");
        style->get_Font()->set_Size(16);
        style->get_Font()->set_Color(System::Drawing::Color::get_Navy());

        // Write a paragraph in that style
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"MyStyle"));
        builder->Writeln(u"Hello world!");

        SharedPtr<Style> firstParagraphStyle = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Style();

        ASPOSE_ASSERT_EQ(style, firstParagraphStyle);

        // Styles can also be removed from the collection like this
        doc->get_Styles()->idx_get(u"MyStyle")->Remove();

        firstParagraphStyle = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Style();

        // Removing the style reverts the styling of the text that was in that style

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
        //ExSummary:Shows how to add a Style to a StyleCollection.
        auto doc = MakeObject<Document>();

        // New documents come with a collection of default styles that can be applied to paragraphs
        SharedPtr<StyleCollection> styles = doc->get_Styles();
        // We can set default parameters for new styles that will be added to the collection from now on
        styles->get_DefaultFont()->set_Name(u"Courier New");
        styles->get_DefaultParagraphFormat()->set_FirstLineIndent(15.0);

        styles->Add(StyleType::Paragraph, u"MyStyle");

        // Styles within the collection can be referenced either by index or name
        // The default font "Courier New" gets automatically applied to any new style added to the collection
        ASSERT_EQ(u"Courier New", styles->idx_get(4)->get_Font()->get_Name());
        ASPOSE_ASSERT_EQ(15.0, styles->idx_get(u"MyStyle")->get_ParagraphFormat()->get_FirstLineIndent());
        //ExEnd
    }

    void ChangeStyleOfTocLevel()
    {
        auto doc = MakeObject<Document>();

        // Retrieve the style used for the first level of the TOC and change the formatting of the style
        doc->get_Styles()->idx_get(StyleIdentifier::Toc1)->get_Font()->set_Bold(true);
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

        // Iterate through all paragraphs formatted using the TOC result based styles; this is any style between TOC and TOC9
        for (auto para : System::IterateOver(doc->GetChildNodes(NodeType::Paragraph, true)->LINQ_OfType<SharedPtr<Paragraph>>()))
        {
            if (para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() >= StyleIdentifier::Toc1 &&
                para->get_ParagraphFormat()->get_Style()->get_StyleIdentifier() <= StyleIdentifier::Toc9)
            {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers
                SharedPtr<TabStop> tab = para->get_ParagraphFormat()->get_TabStops()->idx_get(0);
                // Remove the old tab from the collection
                para->get_ParagraphFormat()->get_TabStops()->RemoveByPosition(tab->get_Position());
                // Insert a new tab using the same properties but at a modified position
                // We could also change the separators used (dots) by passing a different Leader type
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
        //ExSummary:Shows how to copy a style within the same document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // The AddCopy method creates a copy of the specified style and automatically generates a new name for the style, such as "Heading 1_0"
        SharedPtr<Style> newStyle = doc->get_Styles()->AddCopy(doc->get_Styles()->idx_get(u"Heading 1"));
        // You can change the new style name if required as the Style.Name property is read-write
        newStyle->set_Name(u"My Heading 1");
        //ExEnd

        ASSERT_FALSE(newStyle == nullptr);
        ASSERT_EQ(u"My Heading 1", newStyle->get_Name());
        ASSERT_EQ(doc->get_Styles()->idx_get(u"Heading 1")->get_Type(), newStyle->get_Type());
    }

    void CopyStyleDifferentDocument()
    {

        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExSummary:Shows how to import a style from one document into a different document.
        auto dstDoc = MakeObject<Document>();
        auto srcDoc = MakeObject<Document>();

        SharedPtr<Style> srcStyle = srcDoc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle");
        // Change the font of the heading style to red
        srcStyle->get_Font()->set_Color(System::Drawing::Color::get_Red());

        // The AddCopy method can be used to copy a style from a different document
        SharedPtr<Style> newStyle = dstDoc->get_Styles()->AddCopy(srcStyle);

        // The imported style is identical to its source
        ASSERT_EQ(u"MyStyle", newStyle->get_Name());
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), newStyle->get_Font()->get_Color().ToArgb());
        //ExEnd
    }

    void DefaultStyles()
    {
        auto doc = MakeObject<Document>();

        // Add document-wide defaults parameters
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

        // Create a paragraph style and specify some formatting for it
        SharedPtr<Style> style = doc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle1");
        style->get_Font()->set_Size(24);
        style->get_Font()->set_Name(u"Verdana");
        style->get_ParagraphFormat()->set_SpaceAfter(12);

        // Create a list and make sure the paragraphs that use this style will use this list
        style->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::BulletDefault));
        style->get_ListFormat()->set_ListLevelNumber(0);

        // Apply the paragraph style to the current paragraph in the document and add some text
        builder->get_ParagraphFormat()->set_Style(style);
        builder->Writeln(u"Hello World: MyStyle1, bulleted list.");

        // Change to a paragraph style that has no list formatting
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

        // If a style's name has multiple values separated by commas, each one is considered to be a separate alias
        SharedPtr<Style> style = doc->get_Styles()->idx_get(u"MyStyle");
        ASPOSE_ASSERT_EQ(MakeArray<String>({u"MyStyle Alias 1", u"MyStyle Alias 2"}), style->get_Aliases());
        ASSERT_EQ(u"Title", style->get_BaseStyleName());
        ASSERT_EQ(u"MyStyle Char", style->get_LinkedStyleName());

        // A style can be referenced by alias as well as name
        ASPOSE_ASSERT_EQ(style, doc->get_Styles()->idx_get(u"MyStyle Alias 1"));
        //ExEnd
    }
};

} // namespace ApiExamples
