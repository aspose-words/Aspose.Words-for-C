﻿#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Border.h>
#include <Aspose.Words.Cpp/BorderCollection.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/LineStyle.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <drawing/color.h>
#include <system/collections/ienumerator.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
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

class ExBorderCollection : public ApiExampleBase
{
public:
    void GetBordersEnumerator()
    {
        //ExStart
        //ExFor:BorderCollection.GetEnumerator
        //ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Configure the builder's paragraph format settings to create a green wave border on all sides.
        SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Border>>> enumerator = borders->GetEnumerator();
            while (enumerator->MoveNext())
            {
                SharedPtr<Border> border = enumerator->get_Current();
                border->set_Color(System::Drawing::Color::get_Green());
                border->set_LineStyle(LineStyle::Wave);
                border->set_LineWidth(3);
            }
        }

        // Insert a paragraph. Our border settings will determine the appearance of its border.
        builder->Writeln(u"Hello world!");

        doc->Save(ArtifactsDir + u"BorderCollection.GetBordersEnumerator.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"BorderCollection.GetBordersEnumerator.docx");

        for (const auto& border : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
        {
            ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), border->get_Color().ToArgb());
            ASSERT_EQ(LineStyle::Wave, border->get_LineStyle());
            ASPOSE_ASSERT_EQ(3.0, border->get_LineWidth());
        }
    }

    void RemoveAllBorders()
    {
        //ExStart
        //ExFor:BorderCollection.ClearFormatting
        //ExSummary:Shows how to remove all borders from all paragraphs in a document.
        auto doc = MakeObject<Document>(MyDir + u"Borders.docx");

        // The first paragraph of this document has visible borders with these settings.
        SharedPtr<BorderCollection> firstParagraphBorders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();

        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), firstParagraphBorders->get_Color().ToArgb());
        ASSERT_EQ(LineStyle::Single, firstParagraphBorders->get_LineStyle());
        ASPOSE_ASSERT_EQ(3.0, firstParagraphBorders->get_LineWidth());

        // Use the "ClearFormatting" method on each paragraph to remove all borders.
        for (const auto& paragraph : System::IterateOver<Paragraph>(doc->get_FirstSection()->get_Body()->get_Paragraphs()))
        {
            paragraph->get_ParagraphFormat()->get_Borders()->ClearFormatting();

            for (const auto& border : System::IterateOver(paragraph->get_ParagraphFormat()->get_Borders()))
            {
                ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), border->get_Color().ToArgb());
                ASSERT_EQ(LineStyle::None, border->get_LineStyle());
                ASPOSE_ASSERT_EQ(0.0, border->get_LineWidth());
            }
        }

        doc->Save(ArtifactsDir + u"BorderCollection.RemoveAllBorders.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"BorderCollection.RemoveAllBorders.docx");

        for (const auto& border : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
        {
            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), border->get_Color().ToArgb());
            ASSERT_EQ(LineStyle::None, border->get_LineStyle());
            ASPOSE_ASSERT_EQ(0.0, border->get_LineWidth());
        }
    }
};

} // namespace ApiExamples
