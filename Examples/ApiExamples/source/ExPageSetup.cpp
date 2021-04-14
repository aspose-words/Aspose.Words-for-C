// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPageSetup.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Settings;
namespace ApiExamples { namespace gtest_test {

class ExPageSetup : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExPageSetup> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExPageSetup>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExPageSetup> ExPageSetup::s_instance;

TEST_F(ExPageSetup, ClearFormatting)
{
    s_instance->ClearFormatting();
}

using ExPageSetup_DifferentFirstPageHeaderFooter_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPageSetup::DifferentFirstPageHeaderFooter)>::type;

struct ExPageSetup_DifferentFirstPageHeaderFooter : public ExPageSetup,
                                                    public ApiExamples::ExPageSetup,
                                                    public ::testing::WithParamInterface<ExPageSetup_DifferentFirstPageHeaderFooter_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_DifferentFirstPageHeaderFooter, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DifferentFirstPageHeaderFooter(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_DifferentFirstPageHeaderFooter, ::testing::ValuesIn(ExPageSetup_DifferentFirstPageHeaderFooter::TestCases()));

using ExPageSetup_OddAndEvenPagesHeaderFooter_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPageSetup::OddAndEvenPagesHeaderFooter)>::type;

struct ExPageSetup_OddAndEvenPagesHeaderFooter : public ExPageSetup,
                                                 public ApiExamples::ExPageSetup,
                                                 public ::testing::WithParamInterface<ExPageSetup_OddAndEvenPagesHeaderFooter_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_OddAndEvenPagesHeaderFooter, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->OddAndEvenPagesHeaderFooter(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_OddAndEvenPagesHeaderFooter, ::testing::ValuesIn(ExPageSetup_OddAndEvenPagesHeaderFooter::TestCases()));

TEST_F(ExPageSetup, CharactersPerLine)
{
    s_instance->CharactersPerLine();
}

TEST_F(ExPageSetup, LinesPerPage)
{
    s_instance->LinesPerPage();
}

TEST_F(ExPageSetup, SetSectionStart)
{
    s_instance->SetSectionStart();
}

TEST_F(ExPageSetup, PageMargins)
{
    s_instance->PageMargins();
}

TEST_F(ExPageSetup, PaperSizes)
{
    s_instance->PaperSizes();
}

TEST_F(ExPageSetup, ColumnsSameWidth)
{
    s_instance->ColumnsSameWidth();
}

TEST_F(ExPageSetup, CustomColumnWidth)
{
    s_instance->CustomColumnWidth();
}

using ExPageSetup_VerticalLineBetweenColumns_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPageSetup::VerticalLineBetweenColumns)>::type;

struct ExPageSetup_VerticalLineBetweenColumns : public ExPageSetup,
                                                public ApiExamples::ExPageSetup,
                                                public ::testing::WithParamInterface<ExPageSetup_VerticalLineBetweenColumns_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_VerticalLineBetweenColumns, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->VerticalLineBetweenColumns(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_VerticalLineBetweenColumns, ::testing::ValuesIn(ExPageSetup_VerticalLineBetweenColumns::TestCases()));

TEST_F(ExPageSetup, LineNumbers)
{
    s_instance->LineNumbers();
}

TEST_F(ExPageSetup, PageBorderProperties)
{
    s_instance->PageBorderProperties();
}

TEST_F(ExPageSetup, PageBorders)
{
    s_instance->PageBorders();
}

TEST_F(ExPageSetup, PageNumbering)
{
    s_instance->PageNumbering();
}

TEST_F(ExPageSetup, FootnoteOptions)
{
    s_instance->FootnoteOptions_();
}

using ExPageSetup_Bidi_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExPageSetup::Bidi)>::type;

struct ExPageSetup_Bidi : public ExPageSetup, public ApiExamples::ExPageSetup, public ::testing::WithParamInterface<ExPageSetup_Bidi_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExPageSetup_Bidi, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Bidi(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExPageSetup_Bidi, ::testing::ValuesIn(ExPageSetup_Bidi::TestCases()));

TEST_F(ExPageSetup, PageBorder)
{
    s_instance->PageBorder();
}

TEST_F(ExPageSetup, Gutter)
{
    s_instance->Gutter();
}

TEST_F(ExPageSetup, Booklet)
{
    s_instance->Booklet();
}

TEST_F(ExPageSetup, SetTextOrientation)
{
    s_instance->SetTextOrientation();
}

TEST_F(ExPageSetup, SuppressEndnotes)
{
    s_instance->SuppressEndnotes();
}

}} // namespace ApiExamples::gtest_test
