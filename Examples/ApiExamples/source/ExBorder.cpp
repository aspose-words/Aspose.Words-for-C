// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBorder.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExBorder : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExBorder> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExBorder>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExBorder> ExBorder::s_instance;

TEST_F(ExBorder, FontBorder)
{
    s_instance->FontBorder();
}

TEST_F(ExBorder, ParagraphTopBorder)
{
    s_instance->ParagraphTopBorder();
}

TEST_F(ExBorder, ClearFormatting)
{
    s_instance->ClearFormatting();
}

TEST_F(ExBorder, SharedElements)
{
    s_instance->SharedElements();
}

TEST_F(ExBorder, HorizontalBorders)
{
    s_instance->HorizontalBorders();
}

TEST_F(ExBorder, VerticalBorders)
{
    s_instance->VerticalBorders();
}

}} // namespace ApiExamples::gtest_test
