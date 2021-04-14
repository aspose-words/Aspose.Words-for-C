// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported
#include "ExParagraphFormat.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
namespace ApiExamples { namespace gtest_test {

class ExParagraphFormat : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExParagraphFormat> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExParagraphFormat>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExParagraphFormat> ExParagraphFormat::s_instance;

TEST_F(ExParagraphFormat, AsianTypographyProperties)
{
    s_instance->AsianTypographyProperties();
}

using ExParagraphFormat_DropCap_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExParagraphFormat::DropCap)>::type;

struct ExParagraphFormat_DropCap : public ExParagraphFormat,
                                   public ApiExamples::ExParagraphFormat,
                                   public ::testing::WithParamInterface<ExParagraphFormat_DropCap_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(DropCapPosition::Margin),
            std::make_tuple(DropCapPosition::Normal),
            std::make_tuple(DropCapPosition::None),
        };
    }
};

TEST_P(ExParagraphFormat_DropCap, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DropCap(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExParagraphFormat_DropCap, ::testing::ValuesIn(ExParagraphFormat_DropCap::TestCases()));

TEST_F(ExParagraphFormat, LineSpacing)
{
    s_instance->LineSpacing();
}

using ExParagraphFormat_ParagraphSpacingAuto_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExParagraphFormat::ParagraphSpacingAuto)>::type;

struct ExParagraphFormat_ParagraphSpacingAuto : public ExParagraphFormat,
                                                public ApiExamples::ExParagraphFormat,
                                                public ::testing::WithParamInterface<ExParagraphFormat_ParagraphSpacingAuto_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExParagraphFormat_ParagraphSpacingAuto, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ParagraphSpacingAuto(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExParagraphFormat_ParagraphSpacingAuto, ::testing::ValuesIn(ExParagraphFormat_ParagraphSpacingAuto::TestCases()));

using ExParagraphFormat_ParagraphSpacingSameStyle_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExParagraphFormat::ParagraphSpacingSameStyle)>::type;

struct ExParagraphFormat_ParagraphSpacingSameStyle : public ExParagraphFormat,
                                                     public ApiExamples::ExParagraphFormat,
                                                     public ::testing::WithParamInterface<ExParagraphFormat_ParagraphSpacingSameStyle_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExParagraphFormat_ParagraphSpacingSameStyle, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ParagraphSpacingSameStyle(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExParagraphFormat_ParagraphSpacingSameStyle, ::testing::ValuesIn(ExParagraphFormat_ParagraphSpacingSameStyle::TestCases()));

TEST_F(ExParagraphFormat, ParagraphOutlineLevel)
{
    s_instance->ParagraphOutlineLevel();
}

using ExParagraphFormat_PageBreakBefore_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExParagraphFormat::PageBreakBefore)>::type;

struct ExParagraphFormat_PageBreakBefore : public ExParagraphFormat,
                                           public ApiExamples::ExParagraphFormat,
                                           public ::testing::WithParamInterface<ExParagraphFormat_PageBreakBefore_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExParagraphFormat_PageBreakBefore, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PageBreakBefore(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExParagraphFormat_PageBreakBefore, ::testing::ValuesIn(ExParagraphFormat_PageBreakBefore::TestCases()));

using ExParagraphFormat_WidowControl_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExParagraphFormat::WidowControl)>::type;

struct ExParagraphFormat_WidowControl : public ExParagraphFormat,
                                        public ApiExamples::ExParagraphFormat,
                                        public ::testing::WithParamInterface<ExParagraphFormat_WidowControl_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExParagraphFormat_WidowControl, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->WidowControl(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExParagraphFormat_WidowControl, ::testing::ValuesIn(ExParagraphFormat_WidowControl::TestCases()));

TEST_F(ExParagraphFormat, LinesToDrop)
{
    s_instance->LinesToDrop();
}

using ExParagraphFormat_SuppressHyphens_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExParagraphFormat::SuppressHyphens)>::type;

struct ExParagraphFormat_SuppressHyphens : public ExParagraphFormat,
                                           public ApiExamples::ExParagraphFormat,
                                           public ::testing::WithParamInterface<ExParagraphFormat_SuppressHyphens_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExParagraphFormat_SuppressHyphens, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SuppressHyphens(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExParagraphFormat_SuppressHyphens, ::testing::ValuesIn(ExParagraphFormat_SuppressHyphens::TestCases()));

TEST_F(ExParagraphFormat, ParagraphSpacingAndIndents)
{
    s_instance->ParagraphSpacingAndIndents();
}

}} // namespace ApiExamples::gtest_test
