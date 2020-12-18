// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExLayout.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Layout/Model/Callback/PageLayoutCallbackArgs.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
namespace ApiExamples { namespace gtest_test {

class ExLayout : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExLayout> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExLayout>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExLayout> ExLayout::s_instance;

TEST_F(ExLayout, LayoutCollector)
{
    s_instance->LayoutCollector_();
}

TEST_F(ExLayout, LayoutEnumerator)
{
    s_instance->LayoutEnumerator_();
}

TEST_F(ExLayout, PageLayoutCallback)
{
    s_instance->PageLayoutCallback();
}

}} // namespace ApiExamples::gtest_test
