// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExImage.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace ApiExamples { namespace gtest_test {

class ExImage : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExImage> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExImage>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExImage> ExImage::s_instance;

TEST_F(ExImage, FromFile)
{
    s_instance->FromFile();
}

TEST_F(ExImage, FromStream)
{
    s_instance->FromStream();
}

TEST_F(ExImage, SkipMono_FromImage)
{
    RecordProperty("category", "SkipMono");

    s_instance->FromImage();
}

TEST_F(ExImage, CreateFloatingPageCenter)
{
    s_instance->CreateFloatingPageCenter();
}

TEST_F(ExImage, CreateFloatingPositionSize)
{
    s_instance->CreateFloatingPositionSize();
}

TEST_F(ExImage, InsertImageWithHyperlink)
{
    s_instance->InsertImageWithHyperlink();
}

TEST_F(ExImage, CreateLinkedImage)
{
    s_instance->CreateLinkedImage();
}

TEST_F(ExImage, DeleteAllImages)
{
    s_instance->DeleteAllImages();
}

TEST_F(ExImage, DeleteAllImagesPreOrder)
{
    s_instance->DeleteAllImagesPreOrder();
}

TEST_F(ExImage, ScaleImage)
{
    s_instance->ScaleImage();
}

}} // namespace ApiExamples::gtest_test
