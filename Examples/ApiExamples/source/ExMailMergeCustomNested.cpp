// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMailMergeCustomNested.h"

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples { namespace gtest_test {

class ExMailMergeCustomNested : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExMailMergeCustomNested> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExMailMergeCustomNested>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExMailMergeCustomNested> ExMailMergeCustomNested::s_instance;

TEST_F(ExMailMergeCustomNested, CustomDataSource)
{
    s_instance->CustomDataSource();
}

}} // namespace ApiExamples::gtest_test
