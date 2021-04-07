// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExHeaderFooter.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Replacing;
namespace ApiExamples { namespace gtest_test {

class ExHeaderFooter : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExHeaderFooter> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExHeaderFooter>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExHeaderFooter> ExHeaderFooter::s_instance;

TEST_F(ExHeaderFooter, Create)
{
    s_instance->Create();
}

TEST_F(ExHeaderFooter, Link)
{
    s_instance->Link();
}

TEST_F(ExHeaderFooter, RemoveFooters)
{
    s_instance->RemoveFooters();
}

TEST_F(ExHeaderFooter, ExportMode)
{
    s_instance->ExportMode();
}

TEST_F(ExHeaderFooter, ReplaceText)
{
    s_instance->ReplaceText();
}

using ExHeaderFooter_Order_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExHeaderFooter::Order)>::type;

struct ExHeaderFooter_Order : public ExHeaderFooter, public ApiExamples::ExHeaderFooter, public ::testing::WithParamInterface<ExHeaderFooter_Order_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExHeaderFooter_Order, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->Order(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExHeaderFooter_Order, ::testing::ValuesIn(ExHeaderFooter_Order::TestCases()));

TEST_F(ExHeaderFooter, Primer)
{
    s_instance->Primer();
}

}} // namespace ApiExamples::gtest_test
