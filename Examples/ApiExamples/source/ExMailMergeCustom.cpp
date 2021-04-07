// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported
#include "ExMailMergeCustom.h"

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples { namespace gtest_test {

class ExMailMergeCustom : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExMailMergeCustom> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExMailMergeCustom>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExMailMergeCustom> ExMailMergeCustom::s_instance;

TEST_F(ExMailMergeCustom, CustomDataSource)
{
    s_instance->CustomDataSource();
}

TEST_F(ExMailMergeCustom, CustomDataSourceRoot)
{
    s_instance->CustomDataSourceRoot();
}

}} // namespace ApiExamples::gtest_test
