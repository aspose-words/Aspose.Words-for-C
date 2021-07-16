#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/CleanupOptions.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <drawing/color.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace ApiExamples {

class ExCleanupOptions : public ApiExampleBase
{
public:
    void RemoveUnusedResources()
    {
        //ExStart
        //ExFor:Document.Cleanup(CleanupOptions)
        //ExFor:CleanupOptions
        //ExFor:CleanupOptions.UnusedLists
        //ExFor:CleanupOptions.UnusedStyles
        //ExFor:CleanupOptions.UnusedBuiltinStyles
        //ExSummary:Shows how to remove all unused custom styles from a document.
        auto doc = MakeObject<Document>();

        doc->get_Styles()->Add(StyleType::List, u"MyListStyle1");
        doc->get_Styles()->Add(StyleType::List, u"MyListStyle2");
        doc->get_Styles()->Add(StyleType::Character, u"MyParagraphStyle1");
        doc->get_Styles()->Add(StyleType::Character, u"MyParagraphStyle2");

        // Combined with the built-in styles, the document now has eight styles.
        // A custom style is marked as "used" while there is any text within the document
        // formatted in that style. This means that the 4 styles we added are currently unused.
        ASSERT_EQ(8, doc->get_Styles()->get_Count());

        // Apply a custom character style, and then a custom list style. Doing so will mark them as "used".
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_Font()->set_Style(doc->get_Styles()->idx_get(u"MyParagraphStyle1"));
        builder->Writeln(u"Hello world!");

        SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(doc->get_Styles()->idx_get(u"MyListStyle1"));
        builder->get_ListFormat()->set_List(list);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");

        // Now, there is one unused character style and one unused list style.
        // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them.
        auto cleanupOptions = MakeObject<CleanupOptions>();
        cleanupOptions->set_UnusedLists(true);
        cleanupOptions->set_UnusedStyles(true);
        cleanupOptions->set_UnusedBuiltinStyles(true);

        doc->Cleanup(cleanupOptions);

        ASSERT_EQ(4, doc->get_Styles()->get_Count());

        // Removing every node that a custom style is applied to marks it as "unused" again.
        // Rerun the Cleanup method to remove them.
        doc->get_FirstSection()->get_Body()->RemoveAllChildren();
        doc->Cleanup(cleanupOptions);

        ASSERT_EQ(2, doc->get_Styles()->get_Count());
        //ExEnd
    }

    void RemoveDuplicateStyles()
    {
        //ExStart
        //ExFor:CleanupOptions.DuplicateStyle
        //ExSummary:Shows how to remove duplicated styles from the document.
        auto doc = MakeObject<Document>();

        // Add two styles to the document with identical properties,
        // but different names. The second style is considered a duplicate of the first.
        SharedPtr<Style> myStyle = doc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle1");
        myStyle->get_Font()->set_Size(14);
        myStyle->get_Font()->set_Name(u"Courier New");
        myStyle->get_Font()->set_Color(System::Drawing::Color::get_Blue());

        SharedPtr<Style> duplicateStyle = doc->get_Styles()->Add(StyleType::Paragraph, u"MyStyle2");
        duplicateStyle->get_Font()->set_Size(14);
        duplicateStyle->get_Font()->set_Name(u"Courier New");
        duplicateStyle->get_Font()->set_Color(System::Drawing::Color::get_Blue());

        ASSERT_EQ(6, doc->get_Styles()->get_Count());

        // Apply both styles to different paragraphs within the document.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_ParagraphFormat()->set_StyleName(myStyle->get_Name());
        builder->Writeln(u"Hello world!");

        builder->get_ParagraphFormat()->set_StyleName(duplicateStyle->get_Name());
        builder->Writeln(u"Hello again!");

        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASPOSE_ASSERT_EQ(myStyle, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Style());
        ASPOSE_ASSERT_EQ(duplicateStyle, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Style());

        // Configure a CleanOptions object, then call the Cleanup method to substitute all duplicate styles
        // with the original and remove the duplicates from the document.
        auto cleanupOptions = MakeObject<CleanupOptions>();
        cleanupOptions->set_DuplicateStyle(true);

        doc->Cleanup(cleanupOptions);

        ASSERT_EQ(5, doc->get_Styles()->get_Count());
        ASPOSE_ASSERT_EQ(myStyle, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Style());
        ASPOSE_ASSERT_EQ(myStyle, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Style());
        //ExEnd
    }
};

} // namespace ApiExamples
