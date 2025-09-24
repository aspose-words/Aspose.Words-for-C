// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCellFormat.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellMerge.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1346774089u, ::Aspose::Words::ApiExamples::ExCellFormat, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExCellFormat : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExCellFormat> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExCellFormat>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExCellFormat> ExCellFormat::s_instance;

} // namespace gtest_test

void ExCellFormat::VerticalMerge()
{
    //ExStart
    //ExFor:DocumentBuilder.EndRow
    //ExFor:CellMerge
    //ExFor:CellFormat.VerticalMerge
    //ExSummary:Shows how to merge table cells vertically.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a cell into the first column of the first row.
    // This cell will be the first in a range of vertically merged cells.
    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::First);
    builder->Write(u"Text in merged cells.");
    
    // Insert a cell into the second column of the first row, then end the row.
    // Also, configure the builder to disable vertical merging in created cells.
    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::None);
    builder->Write(u"Text in unmerged cell.");
    builder->EndRow();
    
    // Insert a cell into the first column of the second row. 
    // Instead of adding text contents, we will merge this cell with the first cell that we added directly above.
    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::Previous);
    
    // Insert another independent cell in the second column of the second row.
    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::None);
    builder->Write(u"Text in unmerged cell.");
    builder->EndRow();
    builder->EndTable();
    
    doc->Save(get_ArtifactsDir() + u"CellFormat.VerticalMerge.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"CellFormat.VerticalMerge.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::First, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::Previous, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalMerge());
    ASSERT_EQ(u"Text in merged cells.", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim(u'\a'));
    ASSERT_NE(table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText(), table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->GetText());
}

namespace gtest_test
{

TEST_F(ExCellFormat, VerticalMerge)
{
    s_instance->VerticalMerge();
}

} // namespace gtest_test

void ExCellFormat::HorizontalMerge()
{
    //ExStart
    //ExFor:CellMerge
    //ExFor:CellFormat.HorizontalMerge
    //ExSummary:Shows how to merge table cells horizontally.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a cell into the first column of the first row.
    // This cell will be the first in a range of horizontally merged cells.
    builder->InsertCell();
    builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::First);
    builder->Write(u"Text in merged cells.");
    
    // Insert a cell into the second column of the first row. Instead of adding text contents,
    // we will merge this cell with the first cell that we added directly to the left.
    builder->InsertCell();
    builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::Previous);
    builder->EndRow();
    
    // Insert two more unmerged cells to the second row.
    builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::None);
    builder->InsertCell();
    builder->Write(u"Text in unmerged cell.");
    builder->InsertCell();
    builder->Write(u"Text in unmerged cell.");
    builder->EndRow();
    builder->EndTable();
    
    doc->Save(get_ArtifactsDir() + u"CellFormat.HorizontalMerge.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"CellFormat.HorizontalMerge.docx");
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    ASSERT_EQ(1, table->get_Rows()->idx_get(0)->get_Cells()->get_Count());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::None, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_HorizontalMerge());
    ASSERT_EQ(u"Text in merged cells.", table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->GetText().Trim(u'\a'));
}

namespace gtest_test
{

TEST_F(ExCellFormat, HorizontalMerge)
{
    s_instance->HorizontalMerge();
}

} // namespace gtest_test

void ExCellFormat::Padding()
{
    //ExStart
    //ExFor:CellFormat.SetPaddings
    //ExSummary:Shows how to pad the contents of a cell with whitespace.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set a padding distance (in points) between the border and the text contents
    // of each table cell we create with the document builder. 
    builder->get_CellFormat()->SetPaddings(5, 10, 40, 50);
    
    // Create a table with one cell whose contents will have whitespace padding.
    builder->StartTable();
    builder->InsertCell();
    builder->Write(System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") + u"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
    
    doc->Save(get_ArtifactsDir() + u"CellFormat.Padding.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"CellFormat.Padding.docx");
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    System::SharedPtr<Aspose::Words::Tables::Cell> cell = table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0);
    
    ASPOSE_ASSERT_EQ(5, cell->get_CellFormat()->get_LeftPadding());
    ASPOSE_ASSERT_EQ(10, cell->get_CellFormat()->get_TopPadding());
    ASPOSE_ASSERT_EQ(40, cell->get_CellFormat()->get_RightPadding());
    ASPOSE_ASSERT_EQ(50, cell->get_CellFormat()->get_BottomPadding());
}

namespace gtest_test
{

TEST_F(ExCellFormat, Padding)
{
    s_instance->Padding();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
