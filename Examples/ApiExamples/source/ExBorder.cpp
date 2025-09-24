// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBorder.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/enumerator_adapter.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Themes/ThemeColor.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
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


using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Themes;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1586240732u, ::Aspose::Words::ApiExamples::ExBorder, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExBorder : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExBorder> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExBorder>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExBorder> ExBorder::s_instance;

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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->get_Border()->set_Color(System::Drawing::Color::get_Green());
    builder->get_Font()->get_Border()->set_LineWidth(2.5);
    builder->get_Font()->get_Border()->set_LineStyle(Aspose::Words::LineStyle::DashDotStroker);
    
    builder->Write(u"Text surrounded by green border.");
    
    doc->Save(get_ArtifactsDir() + u"Border.FontBorder.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Border.FontBorder.docx");
    System::SharedPtr<Aspose::Words::Border> border = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Border();
    
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
    //ExFor:Border.ThemeColor
    //ExFor:Border.TintAndShade
    //ExFor:Border
    //ExFor:BorderType
    //ExFor:ParagraphFormat.Borders
    //ExSummary:Shows how to insert a paragraph with a top border.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Border> topBorder = builder->get_ParagraphFormat()->get_Borders()->get_Top();
    topBorder->set_LineWidth(4.0);
    topBorder->set_LineStyle(Aspose::Words::LineStyle::DashSmallGap);
    // Set ThemeColor only when LineWidth or LineStyle setted.
    topBorder->set_ThemeColor(Aspose::Words::Themes::ThemeColor::Accent1);
    topBorder->set_TintAndShade(0.25);
    
    builder->Writeln(u"Text with a top border.");
    
    doc->Save(get_ArtifactsDir() + u"Border.ParagraphTopBorder.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Border.ParagraphTopBorder.docx");
    System::SharedPtr<Aspose::Words::Border> border = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()->get_Top();
    
    ASPOSE_ASSERT_EQ(4.0, border->get_LineWidth());
    ASSERT_EQ(Aspose::Words::LineStyle::DashSmallGap, border->get_LineStyle());
    ASSERT_EQ(Aspose::Words::Themes::ThemeColor::Accent1, border->get_ThemeColor());
    ASSERT_NEAR(0.25, border->get_TintAndShade(), 0.01);
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
    //ExFor:Border.IsVisible
    //ExSummary:Shows how to remove borders from a paragraph.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Borders.docx");
    
    // Each paragraph has an individual set of borders.
    // We can access the settings for the appearance of these borders via the paragraph format object.
    System::SharedPtr<Aspose::Words::BorderCollection> borders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();
    
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), borders->idx_get(0)->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(3.0, borders->idx_get(0)->get_LineWidth());
    ASSERT_EQ(Aspose::Words::LineStyle::Single, borders->idx_get(0)->get_LineStyle());
    ASSERT_TRUE(borders->idx_get(0)->get_IsVisible());
    
    // We can remove a border at once by running the ClearFormatting method. 
    // Running this method on every border of a paragraph will remove all its borders.
    for (auto&& border : System::IterateOver(borders))
    {
        border->ClearFormatting();
    }
    
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), borders->idx_get(0)->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(0.0, borders->idx_get(0)->get_LineWidth());
    ASSERT_EQ(Aspose::Words::LineStyle::None, borders->idx_get(0)->get_LineStyle());
    ASSERT_FALSE(borders->idx_get(0)->get_IsVisible());
    
    doc->Save(get_ArtifactsDir() + u"Border.ClearFormatting.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Border.ClearFormatting.docx");
    
    for (auto&& testBorder : System::IterateOver(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders()))
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

void ExBorder::SharedElements()
{
    //ExStart
    //ExFor:Border.Equals(Object)
    //ExFor:Border.Equals(Border)
    //ExFor:Border.GetHashCode
    //ExFor:BorderCollection.Count
    //ExFor:BorderCollection.Equals(BorderCollection)
    //ExFor:BorderCollection.Item(Int32)
    //ExSummary:Shows how border collections can share elements.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Paragraph 1.");
    builder->Write(u"Paragraph 2.");
    
    // Since we used the same border configuration while creating
    // these paragraphs, their border collections share the same elements.
    System::SharedPtr<Aspose::Words::BorderCollection> firstParagraphBorders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();
    System::SharedPtr<Aspose::Words::BorderCollection> secondParagraphBorders = builder->get_CurrentParagraph()->get_ParagraphFormat()->get_Borders();
    ASSERT_EQ(6, firstParagraphBorders->get_Count());
    //ExSkip
    
    for (int32_t i = 0; i < firstParagraphBorders->get_Count(); i++)
    {
        ASSERT_TRUE(System::ObjectExt::Equals(firstParagraphBorders->idx_get(i), secondParagraphBorders->idx_get(i)));
        ASSERT_EQ(System::ObjectExt::GetHashCode(firstParagraphBorders->idx_get(i)), System::ObjectExt::GetHashCode(secondParagraphBorders->idx_get(i)));
        ASSERT_FALSE(firstParagraphBorders->idx_get(i)->get_IsVisible());
    }
    
    for (auto&& border : System::IterateOver(secondParagraphBorders))
    {
        border->set_LineStyle(Aspose::Words::LineStyle::DotDash);
    }
    
    // After changing the line style of the borders in just the second paragraph,
    // the border collections no longer share the same elements.
    for (int32_t i = 0; i < firstParagraphBorders->get_Count(); i++)
    {
        ASSERT_FALSE(System::ObjectExt::Equals(firstParagraphBorders->idx_get(i), secondParagraphBorders->idx_get(i)));
        ASSERT_NE(System::ObjectExt::GetHashCode(firstParagraphBorders->idx_get(i)), System::ObjectExt::GetHashCode(secondParagraphBorders->idx_get(i)));
        
        // Changing the appearance of an empty border makes it visible.
        ASSERT_TRUE(secondParagraphBorders->idx_get(i)->get_IsVisible());
    }
    
    doc->Save(get_ArtifactsDir() + u"Border.SharedElements.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Border.SharedElements.docx");
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    for (auto&& testBorder : System::IterateOver(paragraphs->idx_get(0)->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(Aspose::Words::LineStyle::None, testBorder->get_LineStyle());
    }
    
    for (auto&& testBorder : System::IterateOver(paragraphs->idx_get(1)->get_ParagraphFormat()->get_Borders()))
    {
        ASSERT_EQ(Aspose::Words::LineStyle::DotDash, testBorder->get_LineStyle());
    }
}

namespace gtest_test
{

TEST_F(ExBorder, SharedElements)
{
    s_instance->SharedElements();
}

} // namespace gtest_test

void ExBorder::HorizontalBorders()
{
    //ExStart
    //ExFor:BorderCollection.Horizontal
    //ExSummary:Shows how to apply settings to horizontal borders to a paragraph's format.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a red horizontal border for the paragraph. Any paragraphs created afterwards will inherit these border settings.
    System::SharedPtr<Aspose::Words::BorderCollection> borders = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_ParagraphFormat()->get_Borders();
    borders->get_Horizontal()->set_Color(System::Drawing::Color::get_Red());
    borders->get_Horizontal()->set_LineStyle(Aspose::Words::LineStyle::DashSmallGap);
    borders->get_Horizontal()->set_LineWidth(3);
    
    // Write text to the document without creating a new paragraph afterward.
    // Since there is no paragraph underneath, the horizontal border will not be visible.
    builder->Write(u"Paragraph above horizontal border.");
    
    // Once we add a second paragraph, the border of the first paragraph will become visible.
    builder->InsertParagraph();
    builder->Write(u"Paragraph below horizontal border.");
    
    doc->Save(get_ArtifactsDir() + u"Border.HorizontalBorders.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Border.HorizontalBorders.docx");
    System::SharedPtr<Aspose::Words::ParagraphCollection> paragraphs = doc->get_FirstSection()->get_Body()->get_Paragraphs();
    
    ASSERT_EQ(Aspose::Words::LineStyle::DashSmallGap, paragraphs->idx_get(0)->get_ParagraphFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Horizontal)->get_LineStyle());
    ASSERT_EQ(Aspose::Words::LineStyle::DashSmallGap, paragraphs->idx_get(1)->get_ParagraphFormat()->get_Borders()->idx_get(Aspose::Words::BorderType::Horizontal)->get_LineStyle());
}

namespace gtest_test
{

TEST_F(ExBorder, HorizontalBorders)
{
    s_instance->HorizontalBorders();
}

} // namespace gtest_test

void ExBorder::VerticalBorders()
{
    //ExStart
    //ExFor:BorderCollection.Horizontal
    //ExFor:BorderCollection.Vertical
    //ExFor:Cell.LastParagraph
    //ExSummary:Shows how to apply settings to vertical borders to a table row's format.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a table with red and blue inner borders.
    System::SharedPtr<Aspose::Words::Tables::Table> table = builder->StartTable();
    
    for (int32_t i = 0; i < 3; i++)
    {
        builder->InsertCell();
        builder->Write(System::String::Format(u"Row {0}, Column 1", i + 1));
        builder->InsertCell();
        builder->Write(System::String::Format(u"Row {0}, Column 2", i + 1));
        
        System::SharedPtr<Aspose::Words::Tables::Row> row = builder->EndRow();
        System::SharedPtr<Aspose::Words::BorderCollection> borders = row->get_RowFormat()->get_Borders();
        
        // Adjust the appearance of borders that will appear between rows.
        borders->get_Horizontal()->set_Color(System::Drawing::Color::get_Red());
        borders->get_Horizontal()->set_LineStyle(Aspose::Words::LineStyle::Dot);
        borders->get_Horizontal()->set_LineWidth(2.0);
        
        // Adjust the appearance of borders that will appear between cells.
        borders->get_Vertical()->set_Color(System::Drawing::Color::get_Blue());
        borders->get_Vertical()->set_LineStyle(Aspose::Words::LineStyle::Dot);
        borders->get_Vertical()->set_LineWidth(2.0);
    }
    
    // A row format, and a cell's inner paragraph use different border settings.
    System::SharedPtr<Aspose::Words::Border> border = table->get_FirstRow()->get_FirstCell()->get_LastParagraph()->get_ParagraphFormat()->get_Borders()->get_Vertical();
    
    ASSERT_EQ(System::Drawing::Color::Empty.ToArgb(), border->get_Color().ToArgb());
    ASPOSE_ASSERT_EQ(0.0, border->get_LineWidth());
    ASSERT_EQ(Aspose::Words::LineStyle::None, border->get_LineStyle());
    
    doc->Save(get_ArtifactsDir() + u"Border.VerticalBorders.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Border.VerticalBorders.docx");
    table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    for (auto&& row : System::IterateOver<Aspose::Words::Tables::Row>(table->GetChildNodes(Aspose::Words::NodeType::Row, true)))
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

TEST_F(ExBorder, VerticalBorders)
{
    s_instance->VerticalBorders();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
