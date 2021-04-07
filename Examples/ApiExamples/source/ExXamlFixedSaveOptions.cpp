// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXamlFixedSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExXamlFixedSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExXamlFixedSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExXamlFixedSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExXamlFixedSaveOptions> ExXamlFixedSaveOptions::s_instance;

TEST_F(ExXamlFixedSaveOptions, ResourceFolder)
{
    s_instance->ResourceFolder();
}

}} // namespace ApiExamples::gtest_test
