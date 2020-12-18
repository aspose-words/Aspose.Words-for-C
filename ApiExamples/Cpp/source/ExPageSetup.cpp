// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPageSetup.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <system/string.h>

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExPageSetup : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExPageSetup> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExPageSetup>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExPageSetup> ExPageSetup::s_instance;

TEST_F(ExPageSetup, ClearFormatting)
{
    s_instance->ClearFormatting();
}

TEST_F(ExPageSetup, DifferentHeaders)
{
    s_instance->DifferentHeaders();
}

TEST_F(ExPageSetup, SetSectionStart)
{
    s_instance->SetSectionStart();
}

TEST_F(ExPageSetup, PageMargins)
{
    s_instance->PageMargins();
}

TEST_F(ExPageSetup, ColumnsSameWidth)
{
    s_instance->ColumnsSameWidth();
}

TEST_F(ExPageSetup, CustomColumnWidth)
{
    s_instance->CustomColumnWidth();
}

TEST_F(ExPageSetup, LineNumbers)
{
    s_instance->LineNumbers();
}

TEST_F(ExPageSetup, PageBorderProperties)
{
    s_instance->PageBorderProperties();
}

TEST_F(ExPageSetup, PageBorders)
{
    s_instance->PageBorders();
}

TEST_F(ExPageSetup, PageNumbering)
{
    s_instance->PageNumbering();
}

TEST_F(ExPageSetup, FootnoteOptions)
{
    s_instance->FootnoteOptions_();
}

TEST_F(ExPageSetup, Bidi)
{
    s_instance->Bidi();
}

TEST_F(ExPageSetup, PageBorder)
{
    s_instance->PageBorder();
}

TEST_F(ExPageSetup, Gutter)
{
    s_instance->Gutter();
}

TEST_F(ExPageSetup, Booklet)
{
    s_instance->Booklet();
}

TEST_F(ExPageSetup, SectionTextOrientation)
{
    s_instance->SectionTextOrientation();
}

TEST_F(ExPageSetup, SuppressEndnotes)
{
    s_instance->SuppressEndnotes();
}

}} // namespace ApiExamples::gtest_test
