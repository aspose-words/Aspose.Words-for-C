// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTable.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExTable : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExTable> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExTable>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExTable> ExTable::s_instance;

TEST_F(ExTable, CreateTable)
{
    s_instance->CreateTable();
}

TEST_F(ExTable, Padding)
{
    s_instance->Padding();
}

TEST_F(ExTable, RowCellFormat)
{
    s_instance->RowCellFormat();
}

TEST_F(ExTable, DisplayContentOfTables)
{
    s_instance->DisplayContentOfTables();
}

TEST_F(ExTable, CalculateDepthOfNestedTables)
{
    s_instance->CalculateDepthOfNestedTables();
}

TEST_F(ExTable, EnsureTableMinimum)
{
    s_instance->EnsureTableMinimum();
}

TEST_F(ExTable, EnsureRowMinimum)
{
    s_instance->EnsureRowMinimum();
}

TEST_F(ExTable, EnsureCellMinimum)
{
    s_instance->EnsureCellMinimum();
}

TEST_F(ExTable, SetOutlineBorders)
{
    s_instance->SetOutlineBorders();
}

TEST_F(ExTable, SetBorders)
{
    s_instance->SetBorders();
}

TEST_F(ExTable, RowFormat)
{
    s_instance->RowFormat_();
}

TEST_F(ExTable, CellFormat)
{
    s_instance->CellFormat_();
}

TEST_F(ExTable, GetDistance)
{
    s_instance->GetDistance();
}

TEST_F(ExTable, Borders)
{
    s_instance->Borders();
}

TEST_F(ExTable, ReplaceCellText)
{
    s_instance->ReplaceCellText();
}

using ExTable_RemoveParagraphTextAndMark_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTable::RemoveParagraphTextAndMark)>::type;

struct ExTable_RemoveParagraphTextAndMark : public ExTable,
                                            public ApiExamples::ExTable,
                                            public ::testing::WithParamInterface<ExTable_RemoveParagraphTextAndMark_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExTable_RemoveParagraphTextAndMark, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RemoveParagraphTextAndMark(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTable_RemoveParagraphTextAndMark, ::testing::ValuesIn(ExTable_RemoveParagraphTextAndMark::TestCases()));

TEST_F(ExTable, PrintTableRange)
{
    s_instance->PrintTableRange();
}

TEST_F(ExTable, CloneTable)
{
    s_instance->CloneTable();
}

using ExTable_AllowBreakAcrossPages_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTable::AllowBreakAcrossPages)>::type;

struct ExTable_AllowBreakAcrossPages : public ExTable, public ApiExamples::ExTable, public ::testing::WithParamInterface<ExTable_AllowBreakAcrossPages_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTable_AllowBreakAcrossPages, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AllowBreakAcrossPages(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTable_AllowBreakAcrossPages, ::testing::ValuesIn(ExTable_AllowBreakAcrossPages::TestCases()));

using ExTable_AllowAutoFitOnTable_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTable::AllowAutoFitOnTable)>::type;

struct ExTable_AllowAutoFitOnTable : public ExTable, public ApiExamples::ExTable, public ::testing::WithParamInterface<ExTable_AllowAutoFitOnTable_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTable_AllowAutoFitOnTable, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AllowAutoFitOnTable(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTable_AllowAutoFitOnTable, ::testing::ValuesIn(ExTable_AllowAutoFitOnTable::TestCases()));

TEST_F(ExTable, KeepTableTogether)
{
    s_instance->KeepTableTogether();
}

TEST_F(ExTable, GetIndexOfTableElements)
{
    s_instance->GetIndexOfTableElements();
}

TEST_F(ExTable, GetPreferredWidthTypeAndValue)
{
    s_instance->GetPreferredWidthTypeAndValue();
}

TEST_F(ExTable, CreateNestedTable)
{
    s_instance->CreateNestedTable();
}

TEST_F(ExTable, CheckCellsMerged)
{
    s_instance->CheckCellsMerged();
}

TEST_F(ExTable, MergeCellRange)
{
    s_instance->MergeCellRange();
}

TEST_F(ExTable, CombineTables)
{
    s_instance->CombineTables();
}

TEST_F(ExTable, SplitTable)
{
    s_instance->SplitTable();
}

TEST_F(ExTable, WrapText)
{
    s_instance->WrapText();
}

TEST_F(ExTable, GetFloatingTableProperties)
{
    s_instance->GetFloatingTableProperties();
}

TEST_F(ExTable, ChangeFloatingTableProperties)
{
    s_instance->ChangeFloatingTableProperties();
}

TEST_F(ExTable, TableStyleCreation)
{
    s_instance->TableStyleCreation();
}

TEST_F(ExTable, SetTableAlignment)
{
    s_instance->SetTableAlignment();
}

TEST_F(ExTable, ConditionalStyles)
{
    s_instance->ConditionalStyles();
}

TEST_F(ExTable, ClearTableStyleFormatting)
{
    s_instance->ClearTableStyleFormatting();
}

TEST_F(ExTable, AlternatingRowStyles)
{
    s_instance->AlternatingRowStyles();
}

TEST_F(ExTable, ConvertToHorizontallyMergedCells)
{
    s_instance->ConvertToHorizontallyMergedCells();
}

}} // namespace ApiExamples::gtest_test
