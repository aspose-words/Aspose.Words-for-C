// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCellFormat.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/HeightRule.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellMerge.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "DocumentHelper.h"

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

class ExCellFormat : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExCellFormat> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExCellFormat>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExCellFormat> ExCellFormat::s_instance;

} // namespace gtest_test

void ExCellFormat::VerticalMerge()
{
    //ExStart
    //ExFor:DocumentBuilder.EndRow
    //ExFor:CellMerge
    //ExFor:CellFormat.VerticalMerge
    //ExSummary:Creates a table with two columns with cells merged vertically in the first column.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::First);
    builder->Write(u"Text in merged cells.");

    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::None);
    builder->Write(u"Text in one cell");
    builder->EndRow();

    builder->InsertCell();
    // This cell is vertically merged to the cell above and should be empty.
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::Previous);

    builder->InsertCell();
    builder->get_CellFormat()->set_VerticalMerge(Aspose::Words::Tables::CellMerge::None);
    builder->Write(u"Text in another cell");
    builder->EndRow();
    builder->EndTable();

    doc->Save(ArtifactsDir + u"CellFormat.VerticalMerge.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"CellFormat.VerticalMerge.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::First, table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalMerge());
    ASSERT_EQ(Aspose::Words::Tables::CellMerge::Previous, table->get_Rows()->idx_get(1)->get_Cells()->idx_get(0)->get_CellFormat()->get_VerticalMerge());

    // After the merge both cells still exist, and the one with the VerticalMerge set to "First" overlaps both of them
    // and only that cell contains the shared text
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
    //ExSummary:Creates a table with two rows with cells in the first row horizontally merged.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertCell();
    builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::First);
    builder->Write(u"Text in merged cells.");

    builder->InsertCell();
    // This cell is merged to the previous and should be empty.
    builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::Previous);
    builder->EndRow();

    builder->InsertCell();
    builder->get_CellFormat()->set_HorizontalMerge(Aspose::Words::Tables::CellMerge::None);
    builder->Write(u"Text in one cell.");

    builder->InsertCell();
    builder->Write(u"Text in another cell.");
    builder->EndRow();
    builder->EndTable();

    doc->Save(ArtifactsDir + u"CellFormat.HorizontalMerge.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"CellFormat.HorizontalMerge.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // Compared to the vertical merge, where both cells are still present,
    // the horizontal merge actually removes cells with a HorizontalMerge set to "Previous" if overlapped by ones with "First"
    // Thus the first row that we inserted two cells into now has one, which is a normal cell with a HorizontalMerge of "None"
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

void ExCellFormat::SetCellPaddings()
{
    //ExStart
    //ExFor:CellFormat.SetPaddings
    //ExSummary:Shows how to set paddings to a table cell.
    auto builder = MakeObject<DocumentBuilder>();

    builder->StartTable();
    builder->get_CellFormat()->set_Width(300);
    builder->get_CellFormat()->SetPaddings(5, 10, 40, 50);

    builder->get_RowFormat()->set_HeightRule(Aspose::Words::HeightRule::Exactly);
    builder->get_RowFormat()->set_Height(50);

    builder->InsertCell();
    builder->Write(u"Row 1, Col 1");
    //ExEnd

    SharedPtr<Document> doc = DocumentHelper::SaveOpen(builder->get_Document());

    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    SharedPtr<Cell> cell = table->get_Rows()->idx_get(0)->get_Cells()->idx_get(0);

    ASPOSE_ASSERT_EQ(5, cell->get_CellFormat()->get_LeftPadding());
    ASPOSE_ASSERT_EQ(10, cell->get_CellFormat()->get_TopPadding());
    ASPOSE_ASSERT_EQ(40, cell->get_CellFormat()->get_RightPadding());
    ASPOSE_ASSERT_EQ(50, cell->get_CellFormat()->get_BottomPadding());
}

namespace gtest_test
{

TEST_F(ExCellFormat, SetCellPaddings)
{
    s_instance->SetCellPaddings();
}

} // namespace gtest_test

} // namespace ApiExamples
