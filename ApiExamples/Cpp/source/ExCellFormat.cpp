// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCellFormat.h"

namespace ApiExamples { namespace gtest_test {

class ExCellFormat : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExCellFormat> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExCellFormat>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExCellFormat> ExCellFormat::s_instance;

TEST_F(ExCellFormat, VerticalMerge)
{
    s_instance->VerticalMerge();
}

TEST_F(ExCellFormat, HorizontalMerge)
{
    s_instance->HorizontalMerge();
}

TEST_F(ExCellFormat, Padding)
{
    s_instance->Padding();
}

}} // namespace ApiExamples::gtest_test
