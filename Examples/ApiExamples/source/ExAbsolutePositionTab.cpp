// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExAbsolutePositionTab.h"

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExAbsolutePositionTab : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExAbsolutePositionTab> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExAbsolutePositionTab>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExAbsolutePositionTab> ExAbsolutePositionTab::s_instance;

TEST_F(ExAbsolutePositionTab, DocumentToTxt)
{
    s_instance->DocumentToTxt();
}

}} // namespace ApiExamples::gtest_test
