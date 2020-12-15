// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExImageSaveOptions.h"

namespace ApiExamples { namespace gtest_test {

class ExImageSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExImageSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExImageSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExImageSaveOptions> ExImageSaveOptions::s_instance;

TEST_F(ExImageSaveOptions, Renderer)
{
    s_instance->Renderer();
}

TEST_F(ExImageSaveOptions, SaveSinglePage)
{
    s_instance->SaveSinglePage();
}

TEST_F(ExImageSaveOptions, SkipMono_WindowsMetaFile)
{
    RecordProperty("category", "SkipMono");

    s_instance->WindowsMetaFile();
}

TEST_F(ExImageSaveOptions, BlackAndWhite)
{
    s_instance->BlackAndWhite();
}

TEST_F(ExImageSaveOptions, SkipMono_FloydSteinbergDithering)
{
    RecordProperty("category", "SkipMono");

    s_instance->FloydSteinbergDithering();
}

TEST_F(ExImageSaveOptions, EditImage)
{
    s_instance->EditImage();
}

}} // namespace ApiExamples::gtest_test
