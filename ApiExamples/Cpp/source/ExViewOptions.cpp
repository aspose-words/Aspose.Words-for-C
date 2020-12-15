// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExViewOptions.h"

namespace ApiExamples { namespace gtest_test {

class ExViewOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExViewOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExViewOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExViewOptions> ExViewOptions::s_instance;

TEST_F(ExViewOptions, SetZoom)
{
    s_instance->SetZoom();
}

TEST_F(ExViewOptions, DisplayBackgroundShape)
{
    s_instance->DisplayBackgroundShape();
}

TEST_F(ExViewOptions, DisplayPageBoundaries)
{
    s_instance->DisplayPageBoundaries();
}

using FormsDesign_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExViewOptions::FormsDesign)>::type;

struct FormsDesignVP : public ExViewOptions, public ApiExamples::ExViewOptions, public ::testing::WithParamInterface<FormsDesign_Args>
{
    static std::vector<FormsDesign_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(FormsDesignVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FormsDesign(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExViewOptions, FormsDesignVP, ::testing::ValuesIn(FormsDesignVP::TestCases()));

}} // namespace ApiExamples::gtest_test
