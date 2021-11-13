// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOoxmlSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples { namespace gtest_test {

class ExOoxmlSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExOoxmlSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExOoxmlSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExOoxmlSaveOptions> ExOoxmlSaveOptions::s_instance;

TEST_F(ExOoxmlSaveOptions, Password)
{
    s_instance->Password();
}

TEST_F(ExOoxmlSaveOptions, Iso29500Strict)
{
    s_instance->Iso29500Strict();
}

using ExOoxmlSaveOptions_RestartingDocumentList_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::RestartingDocumentList)>::type;

struct ExOoxmlSaveOptions_RestartingDocumentList : public ExOoxmlSaveOptions,
                                                   public ApiExamples::ExOoxmlSaveOptions,
                                                   public ::testing::WithParamInterface<ExOoxmlSaveOptions_RestartingDocumentList_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_RestartingDocumentList, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RestartingDocumentList(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_RestartingDocumentList, ::testing::ValuesIn(ExOoxmlSaveOptions_RestartingDocumentList::TestCases()));

using ExOoxmlSaveOptions_LastSavedTime_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::LastSavedTime)>::type;

struct ExOoxmlSaveOptions_LastSavedTime : public ExOoxmlSaveOptions,
                                          public ApiExamples::ExOoxmlSaveOptions,
                                          public ::testing::WithParamInterface<ExOoxmlSaveOptions_LastSavedTime_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_LastSavedTime, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LastSavedTime(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_LastSavedTime, ::testing::ValuesIn(ExOoxmlSaveOptions_LastSavedTime::TestCases()));

using ExOoxmlSaveOptions_KeepLegacyControlChars_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::KeepLegacyControlChars)>::type;

struct ExOoxmlSaveOptions_KeepLegacyControlChars : public ExOoxmlSaveOptions,
                                                   public ApiExamples::ExOoxmlSaveOptions,
                                                   public ::testing::WithParamInterface<ExOoxmlSaveOptions_KeepLegacyControlChars_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_KeepLegacyControlChars, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->KeepLegacyControlChars(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_KeepLegacyControlChars, ::testing::ValuesIn(ExOoxmlSaveOptions_KeepLegacyControlChars::TestCases()));

using ExOoxmlSaveOptions_DocumentCompression_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::DocumentCompression)>::type;

struct ExOoxmlSaveOptions_DocumentCompression : public ExOoxmlSaveOptions,
                                                public ApiExamples::ExOoxmlSaveOptions,
                                                public ::testing::WithParamInterface<ExOoxmlSaveOptions_DocumentCompression_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(CompressionLevel::Maximum),
            std::make_tuple(CompressionLevel::Fast),
            std::make_tuple(CompressionLevel::Normal),
            std::make_tuple(CompressionLevel::SuperFast),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_DocumentCompression, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DocumentCompression(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_DocumentCompression, ::testing::ValuesIn(ExOoxmlSaveOptions_DocumentCompression::TestCases()));

TEST_F(ExOoxmlSaveOptions, CheckFileSignatures)
{
    s_instance->CheckFileSignatures();
}

TEST_F(ExOoxmlSaveOptions, ExportGeneratorName)
{
    s_instance->ExportGeneratorName();
}

}} // namespace ApiExamples::gtest_test
