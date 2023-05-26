#include "Working with table styles and formatting.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Tables { namespace gtest_test {

class WorkingWithTableStylesAndFormatting : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Tables::WorkingWithTableStylesAndFormatting> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Tables::WorkingWithTableStylesAndFormatting>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Tables::WorkingWithTableStylesAndFormatting>
    WorkingWithTableStylesAndFormatting::s_instance;

TEST_F(WorkingWithTableStylesAndFormatting, DistanceBetweenTableSurroundingText)
{
    s_instance->DistanceBetweenTableSurroundingText();
}

TEST_F(WorkingWithTableStylesAndFormatting, ApplyOutlineBorder)
{
    s_instance->ApplyOutlineBorder();
}

TEST_F(WorkingWithTableStylesAndFormatting, BuildTableWithBorders)
{
    s_instance->BuildTableWithBorders();
}

TEST_F(WorkingWithTableStylesAndFormatting, ModifyRowFormatting)
{
    s_instance->ModifyRowFormatting();
}

TEST_F(WorkingWithTableStylesAndFormatting, ApplyRowFormatting)
{
    s_instance->ApplyRowFormatting();
}

TEST_F(WorkingWithTableStylesAndFormatting, CellPadding)
{
    s_instance->CellPadding();
}

TEST_F(WorkingWithTableStylesAndFormatting, ModifyCellFormatting)
{
    s_instance->ModifyCellFormatting();
}

TEST_F(WorkingWithTableStylesAndFormatting, FormatTableAndCellWithDifferentBorders)
{
    s_instance->FormatTableAndCellWithDifferentBorders();
}

TEST_F(WorkingWithTableStylesAndFormatting, TableTitleAndDescription)
{
    s_instance->TableTitleAndDescription();
}

TEST_F(WorkingWithTableStylesAndFormatting, AllowCellSpacing)
{
    s_instance->AllowCellSpacing();
}

TEST_F(WorkingWithTableStylesAndFormatting, BuildTableWithStyle)
{
    s_instance->BuildTableWithStyle();
}

TEST_F(WorkingWithTableStylesAndFormatting, ExpandFormattingOnCellsAndRowFromStyle)
{
    s_instance->ExpandFormattingOnCellsAndRowFromStyle();
}

TEST_F(WorkingWithTableStylesAndFormatting, CreateTableStyle)
{
    s_instance->CreateTableStyle();
}

TEST_F(WorkingWithTableStylesAndFormatting, DefineConditionalFormatting)
{
    s_instance->DefineConditionalFormatting();
}

TEST_F(WorkingWithTableStylesAndFormatting, SetTableCellFormatting)
{
    s_instance->SetTableCellFormatting();
}

TEST_F(WorkingWithTableStylesAndFormatting, SetTableRowFormatting)
{
    s_instance->SetTableRowFormatting();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Tables::gtest_test
