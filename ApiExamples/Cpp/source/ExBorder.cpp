// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBorder.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/LineStyle.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderType.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
namespace ApiExamples {

namespace gtest_test
{

class ExBorder : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExBorder> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExBorder>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExBorder> ExBorder::s_instance;

} // namespace gtest_test

void ExBorder::FontBorder()
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
    builder->get_Font()->get_Border()->set_LineStyle(Aspose::Words::LineStyle::DashDotStroker);

    builder->Write(u"Text surrounded by green border.");

    doc->Save(ArtifactsDir + u"Border.FontBorder.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Border.FontBorder.docx");
    SharedPtr<Border> border = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Border();

    ASSERT_EQ(System::Drawing::Color::get_Green().ToArgb(), border->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(2.5, border->get_LineWidth());
    ASSERT_EQ(Aspose::Words::LineStyle::DashDotStroker, border->get_LineStyle());
}

namespace gtest_test
{

TEST_F(ExBorder, FontBorder)
{
    s_instance->FontBorder();
}

} // namespace gtest_test

void ExBorder::ParagraphTopBorder()
{
    //ExStart
    //ExFor:BorderCollection
    //ExFor:Border
    //ExFor:BorderType
    //ExFor:ParagraphFormat.Borders
    //ExSummary:Shows how to insert a paragraph with a top border.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Border> topBorder = builder->get_ParagraphFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Top);
    topBorder->set_Color(System::Drawing::Color::get_Red());
    topBorder->set_LineWidth(4.0);
    topBorder->set_LineStyle(Aspose::Words::LineStyle::DashSmallGap);

    builder->Writeln(u"Text with a red top border.");

    doc->Save(ArtifactsDir + u"Border.ParagraphTopBorder.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Border.ParagraphTopBorder.docx");
    SharedPtr<Border> border = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Top);

    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), border->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(4.0, border->get_LineWidth());
    ASSERT_EQ(Aspose::Words::LineStyle::DashSmallGap, border->get_LineStyle());
}

namespace gtest_test
{

TEST_F(ExBorder, ParagraphTopBorder)
{
    s_instance->ParagraphTopBorder();
}

} // namespace gtest_test

void ExBorder::ClearFormatting()
{
    //ExStart
    //ExFor:Border.ClearFormatting
    //ExSummary:Shows how to remove borders from a paragraph.
    auto doc = MakeObject<Document>(MyDir + u"Borders.docx");

    // Get the first paragraph's collection of borders
    auto builder = MakeObject<DocumentBuilder>(doc);
    SharedPtr<BorderCollection> borders = builder->get_ParagraphFormat()->get_Borders();
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), borders->idx_get(0)->get_Color().ToArgb());
    //ExSkip
    ASPOSE_ASSERT_EQ(3.0, borders->idx_get(0)->get_LineWidth());
    //ExSkip
    ASSERT_EQ(Aspose::Words::LineStyle::Single, borders->idx_get(0)->get_LineStyle());
    //ExSkip

    for (auto border : System::IterateOver(borders))
    {
        border->ClearFormatting();
    }

    builder->get_CurrentParagraph()->get_Runs()->idx_get(0)->set_Text(u"Paragraph with no border");

    doc->Save(ArtifactsDir + u"Border.ClearFormatting.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Border.ClearFormatting.docx");

    for (auto testBorder : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), testBorder->get_Color().ToArgb());
        ASPOSE_ASSERT_EQ(0.0, testBorder->get_LineWidth());
        ASSERT_EQ(Aspose::Words::LineStyle::None, testBorder->get_LineStyle());
    }
}

namespace gtest_test
{

TEST_F(ExBorder, ClearFormatting)
{
    s_instance->ClearFormatting();
}

} // namespace gtest_test

void ExBorder::EqualityCountingAndVisibility()
{
    //ExStart
    //ExFor:Border.Equals(Object)
    //ExFor:Border.Equals(Border)
    //ExFor:Border.GetHashCode
    //ExFor:Border.IsVisible
    //ExFor:BorderCollection.Count
    //ExFor:BorderCollection.Equals(BorderCollection)
    //ExFor:BorderCollection.Item(Int32)
    //ExSummary:Shows the equality of BorderCollections as well counting, visibility of their elements.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->get_CurrentParagraph()->AppendChild(MakeObject<Run>(doc, u"Paragraph 1."));

    SharedPtr<Paragraph> firstParagraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    SharedPtr<BorderCollection> firstParaBorders = firstParagraph->get_ParagraphFormat()->get_Borders();

    builder->InsertParagraph();
    builder->get_CurrentParagraph()->AppendChild(MakeObject<Run>(doc, u"Paragraph 2."));

    SharedPtr<Paragraph> secondParagraph = builder->get_CurrentParagraph();
    SharedPtr<BorderCollection> secondParaBorders = secondParagraph->get_ParagraphFormat()->get_Borders();

    // Two paragraphs have two different BorderCollections, but share the elements that are in from the first paragraph
    for (int i = 0; i < firstParaBorders->get_Count(); i++)
    {
        ASSERT_TRUE(System::ObjectExt::Equals(firstParaBorders->idx_get(i), secondParaBorders->idx_get(i)));
        ASSERT_EQ(System::ObjectExt::GetHashCode(firstParaBorders->idx_get(i)), System::ObjectExt::GetHashCode(secondParaBorders->idx_get(i)));

        // Borders are invisible by default
        ASSERT_FALSE(firstParaBorders->idx_get(i)->get_IsVisible());
    }

    // Each border in the second paragraph collection becomes no longer the same as its counterpart from the first paragraph collection
    // Change all the elements in the second collection to make it completely different from the first
    ASSERT_EQ(6, secondParaBorders->get_Count());
    //ExSkip
    for (auto border : System::IterateOver(secondParaBorders))
    {
        border->set_LineStyle(Aspose::Words::LineStyle::DotDash);
    }

    // Now the BorderCollections both have their own elements
    for (int i = 0; i < firstParaBorders->get_Count(); i++)
    {
        ASSERT_FALSE(System::ObjectExt::Equals(firstParaBorders->idx_get(i), secondParaBorders->idx_get(i)));
        ASSERT_NE(System::ObjectExt::GetHashCode(firstParaBorders->idx_get(i)), System::ObjectExt::GetHashCode(secondParaBorders->idx_get(i)));
        // Changing the line style made the borders visible
        ASSERT_TRUE(secondParaBorders->idx_get(i)->get_IsVisible());
    }

    doc->Save(ArtifactsDir + u"Border.EqualityCountingAndVisibility.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Border.EqualityCountingAndVisibility.docx");
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    for (auto testBorder : System::IterateOver(paragraphs->idx_get(0)->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(Aspose::Words::LineStyle::None, testBorder->get_LineStyle());
    }

    for (auto testBorder : System::IterateOver(paragraphs->idx_get(1)->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(Aspose::Words::LineStyle::DotDash, testBorder->get_LineStyle());
    }
}

namespace gtest_test
{

TEST_F(ExBorder, EqualityCountingAndVisibility)
{
    s_instance->EqualityCountingAndVisibility();
}

} // namespace gtest_test

void ExBorder::VerticalAndHorizontalBorders()
{
    //ExStart
    //ExFor:BorderCollection.Horizontal
    //ExFor:BorderCollection.Vertical
    //ExFor:Cell.LastParagraph
    //ExSummary:Shows the difference between the Horizontal and Vertical properties of BorderCollection.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // A BorderCollection is one of a Paragraph's formatting properties
    SharedPtr<Paragraph> paragraph = doc->get_FirstSection()->get_Body()->get_FirstParagraph();
    SharedPtr<BorderCollection> paragraphBorders = paragraph->get_ParagraphFormat()->get_Borders();

    // paragraphBorders belongs to the first paragraph, but these changes will apply to subsequently created paragraphs
    paragraphBorders->get_Horizontal()->set_Color(System::Drawing::Color::get_Red());
    paragraphBorders->get_Horizontal()->set_LineStyle(Aspose::Words::LineStyle::DashSmallGap);
    paragraphBorders->get_Horizontal()->set_LineWidth(3);

    // Horizontal borders only appear under a paragraph if there's another paragraph under it
    // Right now the first paragraph has no borders
    builder->get_CurrentParagraph()->AppendChild(MakeObject<Run>(doc, u"Paragraph above horizontal border."));

    // Now the first paragraph will have a red dashed line border under it
    // This new second paragraph can have a border too, but only if we add another paragraph underneath it
    builder->InsertParagraph();
    builder->get_CurrentParagraph()->AppendChild(MakeObject<Run>(doc, u"Paragraph below horizontal border."));

    // A table makes use of both vertical and horizontal properties of BorderCollection
    // Both these properties can only affect the inner borders of a table
    auto table = MakeObject<Table>(doc);
    doc->get_FirstSection()->get_Body()->AppendChild(table);

    for (int i = 0; i < 3; i++)
    {
        auto row = MakeObject<Row>(doc);
        SharedPtr<BorderCollection> rowBorders = row->get_RowFormat()->get_Borders();

        // Vertical borders are ones between rows in a table
        rowBorders->get_Horizontal()->set_Color(System::Drawing::Color::get_Red());
        rowBorders->get_Horizontal()->set_LineStyle(Aspose::Words::LineStyle::Dot);
        rowBorders->get_Horizontal()->set_LineWidth(2.0);

        // Vertical borders are ones between cells in a table
        rowBorders->get_Vertical()->set_Color(System::Drawing::Color::get_Blue());
        rowBorders->get_Vertical()->set_LineStyle(Aspose::Words::LineStyle::Dot);
        rowBorders->get_Vertical()->set_LineWidth(2.0);

        // A blue dotted vertical border will appear between cells
        // A red dotted border will appear between rows
        row->AppendChild(MakeObject<Cell>(doc));
        row->get_LastCell()->AppendChild(MakeObject<Paragraph>(doc));
        row->get_LastCell()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Vertical border to the right."));

        row->AppendChild(MakeObject<Cell>(doc));
        row->get_LastCell()->AppendChild(MakeObject<Paragraph>(doc));
        row->get_LastCell()->get_LastParagraph()->AppendChild(MakeObject<Run>(doc, u"Vertical border to the left."));
        table->AppendChild(row);
    }

    doc->Save(ArtifactsDir + u"Border.VerticalAndHorizontalBorders.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Border.VerticalAndHorizontalBorders.docx");
    SharedPtr<ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();

    ASSERT_EQ(Aspose::Words::LineStyle::DashSmallGap, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Horizontal)->get_LineStyle());
    ASSERT_EQ(Aspose::Words::LineStyle::DashSmallGap, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Horizontal)->get_LineStyle());

    auto outTable = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    for (auto row : System::IterateOver<Row>(outTable->GetChildNodes(Aspose::Words::NodeType::Row, true)))
    {
        ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), row->get_RowFormat()->get_Borders()->get_Horizontal()->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::Dot, row->get_RowFormat()->get_Borders()->get_Horizontal()->get_LineStyle());
        ASPOSE_ASSERT_EQ(2.0, row->get_RowFormat()->get_Borders()->get_Horizontal()->get_LineWidth());

        ASSERT_EQ(System::Drawing::Color::get_Blue().ToArgb(), row->get_RowFormat()->get_Borders()->get_Vertical()->get_Color().ToArgb());
        ASSERT_EQ(Aspose::Words::LineStyle::Dot, row->get_RowFormat()->get_Borders()->get_Vertical()->get_LineStyle());
        ASPOSE_ASSERT_EQ(2.0, row->get_RowFormat()->get_Borders()->get_Vertical()->get_LineWidth());
    }
}

namespace gtest_test
{

TEST_F(ExBorder, VerticalAndHorizontalBorders)
{
    s_instance->VerticalAndHorizontalBorders();
}

} // namespace gtest_test

} // namespace ApiExamples
