#include "Working with Tables.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Tables { namespace gtest_test {

class WorkingWithTables : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Tables::WorkingWithTables> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Tables::WorkingWithTables>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Tables::WorkingWithTables> WorkingWithTables::s_instance;

TEST_F(WorkingWithTables, RemoveColumn)
{
    s_instance->RemoveColumn();
}

TEST_F(WorkingWithTables, InsertBlankColumn)
{
    s_instance->InsertBlankColumn();
}

TEST_F(WorkingWithTables, AutoFitTableToContents)
{
    s_instance->AutoFitTableToContents();
}

TEST_F(WorkingWithTables, AutoFitTableToFixedColumnWidths)
{
    s_instance->AutoFitTableToFixedColumnWidths();
}

TEST_F(WorkingWithTables, AutoFitTableToPageWidth)
{
    s_instance->AutoFitTableToPageWidth();
}

TEST_F(WorkingWithTables, CloneCompleteTable)
{
    s_instance->CloneCompleteTable();
}

TEST_F(WorkingWithTables, CloneLastRow)
{
    s_instance->CloneLastRow();
}

TEST_F(WorkingWithTables, FindingIndex)
{
    s_instance->FindingIndex();
}

TEST_F(WorkingWithTables, InsertTableDirectly)
{
    s_instance->InsertTableDirectly();
}

TEST_F(WorkingWithTables, InsertTableFromHtml)
{
    s_instance->InsertTableFromHtml();
}

TEST_F(WorkingWithTables, CreateSimpleTable)
{
    s_instance->CreateSimpleTable();
}

TEST_F(WorkingWithTables, FormattedTable)
{
    s_instance->FormattedTable();
}

TEST_F(WorkingWithTables, NestedTable)
{
    s_instance->NestedTable();
}

TEST_F(WorkingWithTables, CombineRows)
{
    s_instance->CombineRows();
}

TEST_F(WorkingWithTables, SplitTable)
{
    s_instance->SplitTable();
}

TEST_F(WorkingWithTables, RowFormatDisableBreakAcrossPages)
{
    s_instance->RowFormatDisableBreakAcrossPages();
}

TEST_F(WorkingWithTables, KeepTableTogether)
{
    s_instance->KeepTableTogether();
}

TEST_F(WorkingWithTables, CheckCellsMerged)
{
    s_instance->CheckCellsMerged();
}

TEST_F(WorkingWithTables, VerticalMerge)
{
    s_instance->VerticalMerge();
}

TEST_F(WorkingWithTables, HorizontalMerge)
{
    s_instance->HorizontalMerge();
}

TEST_F(WorkingWithTables, MergeCellRange)
{
    s_instance->MergeCellRange();
}

TEST_F(WorkingWithTables, PrintHorizontalAndVerticalMerged)
{
    s_instance->PrintHorizontalAndVerticalMerged();
}

TEST_F(WorkingWithTables, ConvertToHorizontallyMergedCells)
{
    s_instance->ConvertToHorizontallyMergedCells();
}

TEST_F(WorkingWithTables, RepeatRowsOnSubsequentPages)
{
    s_instance->RepeatRowsOnSubsequentPages();
}

TEST_F(WorkingWithTables, AutoFitToPageWidth)
{
    s_instance->AutoFitToPageWidth();
}

TEST_F(WorkingWithTables, PreferredWidthSettings)
{
    s_instance->PreferredWidthSettings();
}

TEST_F(WorkingWithTables, RetrievePreferredWidthType)
{
    s_instance->RetrievePreferredWidthType();
}

TEST_F(WorkingWithTables, GetTablePosition)
{
    s_instance->GetTablePosition();
}

TEST_F(WorkingWithTables, GetFloatingTablePosition)
{
    s_instance->GetFloatingTablePosition();
}

TEST_F(WorkingWithTables, FloatingTablePosition)
{
    s_instance->FloatingTablePosition();
}

TEST_F(WorkingWithTables, SetRelativeHorizontalOrVerticalPosition)
{
    s_instance->SetRelativeHorizontalOrVerticalPosition();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Tables::gtest_test
