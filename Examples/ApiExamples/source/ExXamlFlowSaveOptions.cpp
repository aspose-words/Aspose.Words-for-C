// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXamlFlowSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExXamlFlowSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExXamlFlowSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExXamlFlowSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExXamlFlowSaveOptions> ExXamlFlowSaveOptions::s_instance;

TEST_F(ExXamlFlowSaveOptions, ImageFolder)
{
    s_instance->ImageFolder();
}

}} // namespace ApiExamples::gtest_test
