// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBuilderImages.h"

namespace ApiExamples { namespace gtest_test {

class ExDocumentBuilderImages : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDocumentBuilderImages> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDocumentBuilderImages>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDocumentBuilderImages> ExDocumentBuilderImages::s_instance;

TEST_F(ExDocumentBuilderImages, InsertImageFromStream)
{
    s_instance->InsertImageFromStream();
}

TEST_F(ExDocumentBuilderImages, InsertImageFromFilename)
{
    s_instance->InsertImageFromFilename();
}

TEST_F(ExDocumentBuilderImages, InsertImageFromImageObject)
{
    s_instance->InsertImageFromImageObject();
}

TEST_F(ExDocumentBuilderImages, InsertImageFromByteArray)
{
    s_instance->InsertImageFromByteArray();
}

}} // namespace ApiExamples::gtest_test
