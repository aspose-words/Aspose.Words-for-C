// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExTxtSaveOptions.h"

#include <Aspose.Words.Cpp/Model/Saving/TxtExportHeadersFootersMode.h>

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

using PageBreaks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::PageBreaks)>::type;

struct PageBreaksVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<PageBreaks_Args>
{
    static std::vector<PageBreaks_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(PageBreaksVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PageBreaks(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, PageBreaksVP, ::testing::ValuesIn(PageBreaksVP::TestCases()));

using AddBidiMarks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::AddBidiMarks)>::type;

struct AddBidiMarksVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<AddBidiMarks_Args>
{
    static std::vector<AddBidiMarks_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(AddBidiMarksVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AddBidiMarks(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, AddBidiMarksVP, ::testing::ValuesIn(AddBidiMarksVP::TestCases()));

using ExportHeadersFooters_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::ExportHeadersFooters)>::type;

struct ExportHeadersFootersVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<ExportHeadersFooters_Args>
{
    static std::vector<ExportHeadersFooters_Args> TestCases()
    {
        return {
            std::make_tuple(TxtExportHeadersFootersMode::AllAtEnd),
            std::make_tuple(TxtExportHeadersFootersMode::PrimaryOnly),
            std::make_tuple(TxtExportHeadersFootersMode::None),
        };
    }
};

TEST_P(ExportHeadersFootersVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ExportHeadersFooters(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, ExportHeadersFootersVP, ::testing::ValuesIn(ExportHeadersFootersVP::TestCases()));

TEST_F(ExTxtSaveOptions, TxtListIndentation)
{
    s_instance->TxtListIndentation_();
}

using SimplifyListLabels_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::SimplifyListLabels)>::type;

struct SimplifyListLabelsVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<SimplifyListLabels_Args>
{
    static std::vector<SimplifyListLabels_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(SimplifyListLabelsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SimplifyListLabels(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, SimplifyListLabelsVP, ::testing::ValuesIn(SimplifyListLabelsVP::TestCases()));

TEST_F(ExTxtSaveOptions, ParagraphBreak)
{
    s_instance->ParagraphBreak();
}

TEST_F(ExTxtSaveOptions, Encoding)
{
    s_instance->Encoding_();
}

using PreserveTableLayout_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExTxtSaveOptions::PreserveTableLayout)>::type;

struct PreserveTableLayoutVP : public ExTxtSaveOptions, public ApiExamples::ExTxtSaveOptions, public ::testing::WithParamInterface<PreserveTableLayout_Args>
{
    static std::vector<PreserveTableLayout_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(PreserveTableLayoutVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PreserveTableLayout(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExTxtSaveOptions, PreserveTableLayoutVP, ::testing::ValuesIn(PreserveTableLayoutVP::TestCases()));

TEST_F(ExTxtSaveOptions, UpdateTableLayout)
{
    s_instance->UpdateTableLayout();
}

}} // namespace ApiExamples::gtest_test
