#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <drawing/color.h>
#include <system/enumerator_adapter.h>
#include <system/object_ext.h>
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
using namespace Aspose::Words::Tables;

namespace ApiExamples {

class ExBorder : public ApiExampleBase
{
public:
    void FontBorder()
    {
        //ExStart
        //ExFor:Border
        //ExFor:Border.Color
        //ExFor:Border.LineWidth
        //ExFor:Border.LineStyle
        //ExFor:Font.Border
        //ExFor:LineStyle
        //ExFor:Font
        //ExFor:DocumentBuilder.Font
        //ExFor:DocumentBuilder.Write(String)
        //ExSummary:Shows how to insert a string surrounded by a border into a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->get_Border()->set_Color(System::Drawing::Color::get_Green());
        builder->get_Font()->get_Border()->set_LineWidth(2.5);
        builder->get_Font()->get_Border()->set_LineStyle(LineStyle::DashDotStroker);

        builder->Write(u"Text surrounded by green border.");

        doc->Save(ArtifactsDir + u"Border.FontBorder.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Border.FontBorder.docx");
        SharedPtr<Border> border = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Border();

        ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), border->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(2.5, border->get_LineWidth());
        ASSERT_EQ(LineStyle::DashDotStroker, border->get_LineStyle());
    }

    void ParagraphTopBorder()
    {
        //ExStart
        //ExFor:BorderCollection
        //ExFor:Border
        //ExFor:BorderType
        //ExFor:ParagraphFormat.Borders
        //ExSummary:Shows how to insert a paragraph with a top border.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Border> topBorder = builder->get_ParagraphFormat()->get_Borders()->idx_get(BorderType::Top);
        topBorder->set_Color(System::Drawing::Color::get_Red());
        topBorder->set_LineWidth(4.0);
        topBorder->set_LineStyle(LineStyle::DashSmallGap);

        builder->Writeln(u"Text with a red top border.");

        doc->Save(ArtifactsDir + u"Border.ParagraphTopBorder.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Border.ParagraphTopBorder.docx");
        SharedPtr<Border> border = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()->idx_get(BorderType::Top);

        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), border->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(4.0, border->get_LineWidth());
        ASSERT_EQ(LineStyle::DashSmallGap, border->get_LineStyle());
    }

    void ClearFormatting()
    {
        //ExStart
        //ExFor:Border.ClearFormatting
        //ExFor:Border.IsVisible
        //ExSummary:Shows how to remove borders from a paragraph.
        auto doc = MakeObject<Document>(MyDir + u"Borders.docx");

        // Each paragraph has an individual set of borders.
        // We can access the settings for the appearance of these borders via the paragraph format object.
        SharedPtr<BorderCollection> borders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();

        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), borders->idx_get(0)->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(3.0, borders->idx_get(0)->get_LineWidth());
        ASSERT_EQ(LineStyle::Single, borders->idx_get(0)->get_LineStyle());
        ASSERT_TRUE(borders->idx_get(0)->get_IsVisible());

        // We can remove a border at once by running the ClearFormatting method.
        // Running this method on every border of a paragraph will remove all its borders.
        for (auto border : System::IterateOver(borders))
        {
            border->ClearFormatting();
        }

        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), borders->idx_get(0)->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(0.0, borders->idx_get(0)->get_LineWidth());
        ASSERT_EQ(LineStyle::None, borders->idx_get(0)->get_LineStyle());
        ASSERT_FALSE(borders->idx_get(0)->get_IsVisible());

        doc->Save(ArtifactsDir + u"Border.ClearFormatting.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Border.ClearFormatting.docx");

        for (auto testBorder : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
        {
            ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), testBorder->get_Color().ToArgb());
            ASPOSE_ASSERT_EQ(0.0, testBorder->get_LineWidth());
            ASSERT_EQ(LineStyle::None, testBorder->get_LineStyle());
        }
    }

    void SharedElements()
    {
        //ExStart
        //ExFor:Border.Equals(Object)
        //ExFor:Border.Equals(Border)
        //ExFor:Border.GetHashCode
        //ExFor:BorderCollection.Count
        //ExFor:BorderCollection.Equals(BorderCollection)
        //ExFor:BorderCollection.Item(Int32)
        //ExSummary:Shows how border collections can share elements.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Paragraph 1.");
        builder->Write(u"Paragraph 2.");

        // Since we used the same border configuration while creating
        // these paragraphs, their border collections share the same elements.
        SharedPtr<BorderCollection> firstParagraphBorders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();
        SharedPtr<BorderCollection> secondParagraphBorders = builder->get_CurrentParagraph()->get_ParagraphFormat()->get_Borders();
        ASSERT_EQ(6, firstParagraphBorders->get_Count());
        //ExSkip

        for (int i = 0; i < firstParagraphBorders->get_Count(); i++)
        {
            ASSERT_TRUE(System::ObjectExt::Equals(firstParagraphBorders->idx_get(i), secondParagraphBorders->idx_get(i)));
            ASSERT_EQ(System::ObjectExt::GetHashCode(firstParagraphBorders->idx_get(i)), System::ObjectExt::GetHashCode(secondParagraphBorders->idx_get(i)));
            ASSERT_FALSE(firstParagraphBorders->idx_get(i)->get_IsVisible());
        }

        for (auto border : System::IterateOver(secondParagraphBorders))
        {
            border->set_LineStyle(LineStyle::DotDash);
        }

        // After changing the line style of the borders in just the second paragraph,
        // the border collections no longer share the same elements.
        for (int i = 0; i < firstParagraphBorders->get_Count(); i++)
        {
            ASSERT_FALSE(System::ObjectExt::Equals(firstParagraphBorders->idx_get(i), secondParagraphBorders->idx_get(i)));
            ASSERT_NE(System::ObjectExt::GetHashCode(firstParagraphBorders->idx_get(i)), System::ObjectExt::GetHashCode(secondParagraphBorders->idx_get(i)));

            // Changing the appearance of an empty border makes it visible.
            ASSERT_TRUE(secondParagraphBorders->idx_get(i)->get_IsVisible());
        }

        doc->Save(ArtifactsDir + u"Border.SharedElements.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Border.SharedElements.docx");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        for (auto testBorder : System::IterateOver(paragraphs->idx_get(0)->get_ParagraphFormat()->get_Borders()))
        {
            ASSERT_EQ(LineStyle::None, testBorder->get_LineStyle());
        }

        for (auto testBorder : System::IterateOver(paragraphs->idx_get(1)->get_ParagraphFormat()->get_Borders()))
        {
            ASSERT_EQ(LineStyle::DotDash, testBorder->get_LineStyle());
        }
    }

    void HorizontalBorders()
    {
        //ExStart
        //ExFor:BorderCollection.Horizontal
        //ExSummary:Shows how to apply settings to horizontal borders to a paragraph's format.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a red horizontal border for the paragraph. Any paragraphs created afterwards will inherit these border settings.
        SharedPtr<BorderCollection> borders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();
        borders->get_Horizontal()->set_Color(System::Drawing::Color::get_Red());
        borders->get_Horizontal()->set_LineStyle(LineStyle::DashSmallGap);
        borders->get_Horizontal()->set_LineWidth(3);

        // Write text to the document without creating a new paragraph afterward.
        // Since there is no paragraph underneath, the horizontal border will not be visible.
        builder->Write(u"Paragraph above horizontal border.");

        // Once we add a second paragraph, the border of the first paragraph will become visible.
        builder->InsertParagraph();
        builder->Write(u"Paragraph below horizontal border.");

        doc->Save(ArtifactsDir + u"Border.HorizontalBorders.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Border.HorizontalBorders.docx");
        SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

        ASSERT_EQ(LineStyle::DashSmallGap, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Borders()->idx_get(BorderType::Horizontal)->get_LineStyle());
        ASSERT_EQ(LineStyle::DashSmallGap, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Borders()->idx_get(BorderType::Horizontal)->get_LineStyle());
    }

    void VerticalBorders()
    {
        //ExStart
        //ExFor:BorderCollection.Horizontal
        //ExFor:BorderCollection.Vertical
        //ExFor:Cell.LastParagraph
        //ExSummary:Shows how to apply settings to vertical borders to a table row's format.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a table with red and blue inner borders.
        SharedPtr<Table> table = builder->StartTable();

        for (int i = 0; i < 3; i++)
        {
            builder->InsertCell();
            builder->Write(String::Format(u"Row {0}, Column 1", i + 1));
            builder->InsertCell();
            builder->Write(String::Format(u"Row {0}, Column 2", i + 1));

            SharedPtr<Row> row = builder->EndRow();
            SharedPtr<BorderCollection> borders = row->get_RowFormat()->get_Borders();

            // Adjust the appearance of borders that will appear between rows.
            borders->get_Horizontal()->set_Color(System::Drawing::Color::get_Red());
            borders->get_Horizontal()->set_LineStyle(LineStyle::Dot);
            borders->get_Horizontal()->set_LineWidth(2.0);

            // Adjust the appearance of borders that will appear between cells.
            borders->get_Vertical()->set_Color(System::Drawing::Color::get_Blue());
            borders->get_Vertical()->set_LineStyle(LineStyle::Dot);
            borders->get_Vertical()->set_LineWidth(2.0);
        }

        // A row format, and a cell's inner paragraph use different border settings.
        SharedPtr<Border> border = table->get_FirstRow()->get_FirstCell()->get_LastParagraph()->get_ParagraphFormat()->get_Borders()->get_Vertical();

        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), border->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(0.0, border->get_LineWidth());
        ASSERT_EQ(LineStyle::None, border->get_LineStyle());

        doc->Save(ArtifactsDir + u"Border.VerticalBorders.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Border.VerticalBorders.docx");
        table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);

        for (auto row : System::IterateOver<Row>(table->GetChildNodes(NodeType::Row, true)))
        {
            ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), row->get_RowFormat()->get_Borders()->get_Horizontal()->get_Color().ToArgb());
            ASSERT_EQ(LineStyle::Dot, row->get_RowFormat()->get_Borders()->get_Horizontal()->get_LineStyle());
            ASPOSE_ASSERT_EQ(2.0, row->get_RowFormat()->get_Borders()->get_Horizontal()->get_LineWidth());

            ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), row->get_RowFormat()->get_Borders()->get_Vertical()->get_Color().ToArgb());
            ASSERT_EQ(LineStyle::Dot, row->get_RowFormat()->get_Borders()->get_Vertical()->get_LineStyle());
            ASPOSE_ASSERT_EQ(2.0, row->get_RowFormat()->get_Borders()->get_Vertical()->get_LineWidth());
        }
    }
};

} // namespace ApiExamples
