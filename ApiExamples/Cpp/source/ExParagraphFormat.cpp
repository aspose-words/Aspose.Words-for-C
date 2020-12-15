// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExParagraphFormat.h"

namespace ApiExamples { namespace gtest_test {

class ExParagraphFormat : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExParagraphFormat> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

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

TEST_F(ExParagraphFormat, DropCap)
{
    s_instance->DropCap();
}

TEST_F(ExParagraphFormat, LineSpacing)
{
    s_instance->LineSpacing();
}

TEST_F(ExParagraphFormat, ParagraphSpacing)
{
    s_instance->ParagraphSpacing();
}

TEST_F(ExParagraphFormat, ParagraphOutlineLevel)
{
    s_instance->ParagraphOutlineLevel();
}

TEST_F(ExParagraphFormat, PageBreakBefore)
{
    s_instance->PageBreakBefore();
}

TEST_F(ExParagraphFormat, WidowControl)
{
    s_instance->WidowControl();
}

TEST_F(ExParagraphFormat, LinesToDrop)
{
    s_instance->LinesToDrop();
}

TEST_F(ExParagraphFormat, SuppressHyphens)
{
    s_instance->SuppressHyphens();
}

TEST_F(ExParagraphFormat, SnapToGrid)
{
    s_instance->SnapToGrid();
}

}} // namespace ApiExamples::gtest_test
