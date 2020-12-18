// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExLoadOptions.h"

#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingArgs.h>
#include <system/collections/list.h>
#include <system/shared_ptr.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
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

using ConvertShapeToOfficeMath_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExLoadOptions::ConvertShapeToOfficeMath)>::type;

struct ConvertShapeToOfficeMathVP : public ExLoadOptions, public ApiExamples::ExLoadOptions, public ::testing::WithParamInterface<ConvertShapeToOfficeMath_Args>
{
    static std::vector<ConvertShapeToOfficeMath_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ConvertShapeToOfficeMathVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ConvertShapeToOfficeMath(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExLoadOptions, ConvertShapeToOfficeMathVP, ::testing::ValuesIn(ConvertShapeToOfficeMathVP::TestCases()));

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

}} // namespace ApiExamples::gtest_test
