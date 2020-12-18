// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHtmlLoadOptions.h"

#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <system/collections/list.h>
#include <system/shared_ptr.h>

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExHtmlLoadOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExHtmlLoadOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExHtmlLoadOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExHtmlLoadOptions> ExHtmlLoadOptions::s_instance;

using SupportVml_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHtmlLoadOptions::SupportVml)>::type;

struct SupportVmlVP : public ExHtmlLoadOptions, public ApiExamples::ExHtmlLoadOptions, public ::testing::WithParamInterface<SupportVml_Args>
{
    static std::vector<SupportVml_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(SupportVmlVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SupportVml(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExHtmlLoadOptions, SupportVmlVP, ::testing::ValuesIn(SupportVmlVP::TestCases()));

TEST_F(ExHtmlLoadOptions, WebRequestTimeout)
{
    s_instance->WebRequestTimeout();
}

TEST_F(ExHtmlLoadOptions, EncryptedHtml)
{
    s_instance->EncryptedHtml();
}

TEST_F(ExHtmlLoadOptions, BaseUri)
{
    s_instance->BaseUri();
}

TEST_F(ExHtmlLoadOptions, GetSelectAsSdt)
{
    s_instance->GetSelectAsSdt();
}

TEST_F(ExHtmlLoadOptions, GetInputAsFormField)
{
    s_instance->GetInputAsFormField();
}

}} // namespace ApiExamples::gtest_test
