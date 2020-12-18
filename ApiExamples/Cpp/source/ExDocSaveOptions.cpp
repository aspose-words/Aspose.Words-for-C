// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocSaveOptions.h"

namespace ApiExamples { namespace gtest_test {

class ExDocSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDocSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDocSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDocSaveOptions> ExDocSaveOptions::s_instance;

TEST_F(ExDocSaveOptions, SaveAsDoc)
{
    s_instance->SaveAsDoc();
}

TEST_F(ExDocSaveOptions, TempFolder)
{
    s_instance->TempFolder();
}

TEST_F(ExDocSaveOptions, PictureBullets)
{
    s_instance->PictureBullets();
}

using UpdateLastPrintedProperty_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocSaveOptions::UpdateLastPrintedProperty)>::type;

struct UpdateLastPrintedPropertyVP : public ExDocSaveOptions,
                                     public ApiExamples::ExDocSaveOptions,
                                     public ::testing::WithParamInterface<UpdateLastPrintedProperty_Args>
{
    static std::vector<UpdateLastPrintedProperty_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(UpdateLastPrintedPropertyVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateLastPrintedProperty(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocSaveOptions, UpdateLastPrintedPropertyVP, ::testing::ValuesIn(UpdateLastPrintedPropertyVP::TestCases()));

using AlwaysCompressMetafiles_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocSaveOptions::AlwaysCompressMetafiles)>::type;

struct AlwaysCompressMetafilesVP : public ExDocSaveOptions,
                                   public ApiExamples::ExDocSaveOptions,
                                   public ::testing::WithParamInterface<AlwaysCompressMetafiles_Args>
{
    static std::vector<AlwaysCompressMetafiles_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(AlwaysCompressMetafilesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AlwaysCompressMetafiles(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocSaveOptions, AlwaysCompressMetafilesVP, ::testing::ValuesIn(AlwaysCompressMetafilesVP::TestCases()));

}} // namespace ApiExamples::gtest_test
