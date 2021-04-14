// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExStyles.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;
namespace ApiExamples { namespace gtest_test {

class ExStyles : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExStyles> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExStyles>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExStyles> ExStyles::s_instance;

TEST_F(ExStyles, Styles)
{
    s_instance->Styles();
}

TEST_F(ExStyles, CreateStyle)
{
    s_instance->CreateStyle();
}

TEST_F(ExStyles, StyleCollection)
{
    s_instance->StyleCollection_();
}

TEST_F(ExStyles, RemoveStylesFromStyleGallery)
{
    s_instance->RemoveStylesFromStyleGallery();
}

TEST_F(ExStyles, ChangeTocsTabStops)
{
    s_instance->ChangeTocsTabStops();
}

TEST_F(ExStyles, CopyStyleSameDocument)
{
    s_instance->CopyStyleSameDocument();
}

TEST_F(ExStyles, CopyStyleDifferentDocument)
{
    s_instance->CopyStyleDifferentDocument();
}

TEST_F(ExStyles, DefaultStyles)
{
    s_instance->DefaultStyles();
}

TEST_F(ExStyles, ParagraphStyleBulletedList)
{
    s_instance->ParagraphStyleBulletedList();
}

TEST_F(ExStyles, StyleAliases)
{
    s_instance->StyleAliases();
}

}} // namespace ApiExamples::gtest_test
