// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExTxtSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExTxtSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExTxtSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExTxtSaveOptions> ExTxtSaveOptions::s_instance;

using ExTxtSaveOptions_PageBreaks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::PageBreaks)>::type;

struct ExTxtSaveOptions_PageBreaks : public ExTxtSaveOptions,
                                     public ApiExamples::ExTxtSaveOptions,
                                     public ::testing::WithParamInterface<ExTxtSaveOptions_PageBreaks_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_PageBreaks, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PageBreaks(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_PageBreaks, ::testing::ValuesIn(ExTxtSaveOptions_PageBreaks::TestCases()));

using ExTxtSaveOptions_AddBidiMarks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::AddBidiMarks)>::type;

struct ExTxtSaveOptions_AddBidiMarks : public ExTxtSaveOptions,
                                       public ApiExamples::ExTxtSaveOptions,
                                       public ::testing::WithParamInterface<ExTxtSaveOptions_AddBidiMarks_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_AddBidiMarks, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AddBidiMarks(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_AddBidiMarks, ::testing::ValuesIn(ExTxtSaveOptions_AddBidiMarks::TestCases()));

using ExTxtSaveOptions_ExportHeadersFooters_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::ExportHeadersFooters)>::type;

struct ExTxtSaveOptions_ExportHeadersFooters : public ExTxtSaveOptions,
                                               public ApiExamples::ExTxtSaveOptions,
                                               public ::testing::WithParamInterface<ExTxtSaveOptions_ExportHeadersFooters_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(TxtExportHeadersFootersMode::AllAtEnd),
            std::make_tuple(TxtExportHeadersFootersMode::PrimaryOnly),
            std::make_tuple(TxtExportHeadersFootersMode::None),
        };
    }
};

TEST_P(ExTxtSaveOptions_ExportHeadersFooters, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportHeadersFooters(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_ExportHeadersFooters, ::testing::ValuesIn(ExTxtSaveOptions_ExportHeadersFooters::TestCases()));

TEST_F(ExTxtSaveOptions, TxtListIndentation)
{
    s_instance->TxtListIndentation_();
}

using ExTxtSaveOptions_SimplifyListLabels_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::SimplifyListLabels)>::type;

struct ExTxtSaveOptions_SimplifyListLabels : public ExTxtSaveOptions,
                                             public ApiExamples::ExTxtSaveOptions,
                                             public ::testing::WithParamInterface<ExTxtSaveOptions_SimplifyListLabels_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_SimplifyListLabels, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SimplifyListLabels(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_SimplifyListLabels, ::testing::ValuesIn(ExTxtSaveOptions_SimplifyListLabels::TestCases()));

TEST_F(ExTxtSaveOptions, ParagraphBreak)
{
    s_instance->ParagraphBreak();
}

using ExTxtSaveOptions_PreserveTableLayout_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::PreserveTableLayout)>::type;

struct ExTxtSaveOptions_PreserveTableLayout : public ExTxtSaveOptions,
                                              public ApiExamples::ExTxtSaveOptions,
                                              public ::testing::WithParamInterface<ExTxtSaveOptions_PreserveTableLayout_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExTxtSaveOptions_PreserveTableLayout, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PreserveTableLayout(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExTxtSaveOptions_PreserveTableLayout, ::testing::ValuesIn(ExTxtSaveOptions_PreserveTableLayout::TestCases()));

TEST_F(ExTxtSaveOptions, MaxCharactersPerLine)
{
    s_instance->MaxCharactersPerLine();
}

}} // namespace ApiExamples::gtest_test
