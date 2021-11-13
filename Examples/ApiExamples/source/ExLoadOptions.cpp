// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported
#include "ExLoadOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples { namespace gtest_test {

class ExLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExLoadOptions> ExLoadOptions::s_instance;

TEST_F(ExLoadOptions, LoadOptionsCallback)
{
    s_instance->LoadOptionsCallback();
}

using ExLoadOptions_ConvertShapeToOfficeMath_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExLoadOptions::ConvertShapeToOfficeMath)>::type;

struct ExLoadOptions_ConvertShapeToOfficeMath : public ExLoadOptions,
                                                public ApiExamples::ExLoadOptions,
                                                public ::testing::WithParamInterface<ExLoadOptions_ConvertShapeToOfficeMath_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExLoadOptions_ConvertShapeToOfficeMath, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ConvertShapeToOfficeMath(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExLoadOptions_ConvertShapeToOfficeMath, ::testing::ValuesIn(ExLoadOptions_ConvertShapeToOfficeMath::TestCases()));

TEST_F(ExLoadOptions, SetEncoding)
{
    s_instance->SetEncoding();
}

TEST_F(ExLoadOptions, FontSettings)
{
    s_instance->FontSettings_();
}

TEST_F(ExLoadOptions, LoadOptionsMswVersion)
{
    s_instance->LoadOptionsMswVersion();
}

TEST_F(ExLoadOptions, LoadOptionsWarningCallback)
{
    s_instance->LoadOptionsWarningCallback();
}

TEST_F(ExLoadOptions, TempFolder)
{
    s_instance->TempFolder();
}

TEST_F(ExLoadOptions, AddEditingLanguage)
{
    s_instance->AddEditingLanguage();
}

TEST_F(ExLoadOptions, SetEditingLanguageAsDefault)
{
    s_instance->SetEditingLanguageAsDefault();
}

TEST_F(ExLoadOptions, ConvertMetafilesToPng)
{
    s_instance->ConvertMetafilesToPng();
}

TEST_F(ExLoadOptions, OpenChmFile)
{
    s_instance->OpenChmFile();
}

using ExLoadOptions_FlatOpcXmlMappingOnly_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExLoadOptions::FlatOpcXmlMappingOnly)>::type;

struct ExLoadOptions_FlatOpcXmlMappingOnly : public ExLoadOptions,
                                             public ApiExamples::ExLoadOptions,
                                             public ::testing::WithParamInterface<ExLoadOptions_FlatOpcXmlMappingOnly_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExLoadOptions_FlatOpcXmlMappingOnly, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FlatOpcXmlMappingOnly(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExLoadOptions_FlatOpcXmlMappingOnly, ::testing::ValuesIn(ExLoadOptions_FlatOpcXmlMappingOnly::TestCases()));

}} // namespace ApiExamples::gtest_test
