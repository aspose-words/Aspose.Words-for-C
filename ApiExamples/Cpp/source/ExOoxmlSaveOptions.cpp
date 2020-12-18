// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOoxmlSaveOptions.h"

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

using RestartingDocumentList_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::RestartingDocumentList)>::type;

struct RestartingDocumentListVP : public ExOoxmlSaveOptions,
                                  public ApiExamples::ExOoxmlSaveOptions,
                                  public ::testing::WithParamInterface<RestartingDocumentList_Args>
{
    static std::vector<RestartingDocumentList_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(RestartingDocumentListVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RestartingDocumentList(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExOoxmlSaveOptions, RestartingDocumentListVP, ::testing::ValuesIn(RestartingDocumentListVP::TestCases()));

TEST_F(ExOoxmlSaveOptions, UpdatingLastSavedTimeDocument)
{
    s_instance->UpdatingLastSavedTimeDocument();
}

using KeepLegacyControlChars_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::KeepLegacyControlChars)>::type;

struct KeepLegacyControlCharsVP : public ExOoxmlSaveOptions,
                                  public ApiExamples::ExOoxmlSaveOptions,
                                  public ::testing::WithParamInterface<KeepLegacyControlChars_Args>
{
    static std::vector<KeepLegacyControlChars_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(KeepLegacyControlCharsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->KeepLegacyControlChars(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExOoxmlSaveOptions, KeepLegacyControlCharsVP, ::testing::ValuesIn(KeepLegacyControlCharsVP::TestCases()));

TEST_F(ExOoxmlSaveOptions, DocumentCompression)
{
    s_instance->DocumentCompression();
}

TEST_F(ExOoxmlSaveOptions, CheckFileSignatures)
{
    s_instance->CheckFileSignatures();
}

}} // namespace ApiExamples::gtest_test
