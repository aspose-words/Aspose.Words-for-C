// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocSaveOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
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

using ExDocSaveOptions_UpdateLastPrintedProperty_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocSaveOptions::UpdateLastPrintedProperty)>::type;

struct ExDocSaveOptions_UpdateLastPrintedProperty : public ExDocSaveOptions,
                                                    public ApiExamples::ExDocSaveOptions,
                                                    public ::testing::WithParamInterface<ExDocSaveOptions_UpdateLastPrintedProperty_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocSaveOptions_UpdateLastPrintedProperty, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateLastPrintedProperty(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocSaveOptions_UpdateLastPrintedProperty, ::testing::ValuesIn(ExDocSaveOptions_UpdateLastPrintedProperty::TestCases()));

using ExDocSaveOptions_UpdateCreatedTimeProperty_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocSaveOptions::UpdateCreatedTimeProperty)>::type;

struct ExDocSaveOptions_UpdateCreatedTimeProperty : public ExDocSaveOptions,
                                                    public ApiExamples::ExDocSaveOptions,
                                                    public ::testing::WithParamInterface<ExDocSaveOptions_UpdateCreatedTimeProperty_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocSaveOptions_UpdateCreatedTimeProperty, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateCreatedTimeProperty(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocSaveOptions_UpdateCreatedTimeProperty, ::testing::ValuesIn(ExDocSaveOptions_UpdateCreatedTimeProperty::TestCases()));

using ExDocSaveOptions_AlwaysCompressMetafiles_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocSaveOptions::AlwaysCompressMetafiles)>::type;

struct ExDocSaveOptions_AlwaysCompressMetafiles : public ExDocSaveOptions,
                                                  public ApiExamples::ExDocSaveOptions,
                                                  public ::testing::WithParamInterface<ExDocSaveOptions_AlwaysCompressMetafiles_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocSaveOptions_AlwaysCompressMetafiles, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->AlwaysCompressMetafiles(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocSaveOptions_AlwaysCompressMetafiles, ::testing::ValuesIn(ExDocSaveOptions_AlwaysCompressMetafiles::TestCases()));

}} // namespace ApiExamples::gtest_test
