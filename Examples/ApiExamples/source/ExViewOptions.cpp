// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExViewOptions.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Settings;
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

TEST_F(ExViewOptions, SetZoomPercentage)
{
    s_instance->SetZoomPercentage();
}

using ExViewOptions_SetZoomType_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExViewOptions::SetZoomType)>::type;

struct ExViewOptions_SetZoomType : public ExViewOptions, public ApiExamples::ExViewOptions, public ::testing::WithParamInterface<ExViewOptions_SetZoomType_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(ZoomType::PageWidth),
            std::make_tuple(ZoomType::FullPage),
            std::make_tuple(ZoomType::TextFit),
        };
    }
};

TEST_P(ExViewOptions_SetZoomType, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SetZoomType(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_SetZoomType, ::testing::ValuesIn(ExViewOptions_SetZoomType::TestCases()));

using ExViewOptions_DisplayBackgroundShape_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExViewOptions::DisplayBackgroundShape)>::type;

struct ExViewOptions_DisplayBackgroundShape : public ExViewOptions,
                                              public ApiExamples::ExViewOptions,
                                              public ::testing::WithParamInterface<ExViewOptions_DisplayBackgroundShape_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExViewOptions_DisplayBackgroundShape, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DisplayBackgroundShape(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_DisplayBackgroundShape, ::testing::ValuesIn(ExViewOptions_DisplayBackgroundShape::TestCases()));

using ExViewOptions_DisplayPageBoundaries_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExViewOptions::DisplayPageBoundaries)>::type;

struct ExViewOptions_DisplayPageBoundaries : public ExViewOptions,
                                             public ApiExamples::ExViewOptions,
                                             public ::testing::WithParamInterface<ExViewOptions_DisplayPageBoundaries_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExViewOptions_DisplayPageBoundaries, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DisplayPageBoundaries(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_DisplayPageBoundaries, ::testing::ValuesIn(ExViewOptions_DisplayPageBoundaries::TestCases()));

using ExViewOptions_FormsDesign_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExViewOptions::FormsDesign)>::type;

struct ExViewOptions_FormsDesign : public ExViewOptions, public ApiExamples::ExViewOptions, public ::testing::WithParamInterface<ExViewOptions_FormsDesign_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExViewOptions_FormsDesign, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FormsDesign(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExViewOptions_FormsDesign, ::testing::ValuesIn(ExViewOptions_FormsDesign::TestCases()));

}} // namespace ApiExamples::gtest_test
