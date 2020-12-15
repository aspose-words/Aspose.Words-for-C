// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDrawing.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Drawing/GroupShape.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <system/shared_ptr.h>
#include <system/string.h>
#include <system/text/string_builder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace ApiExamples { namespace gtest_test {

class ExDrawing : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDrawing> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDrawing>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDrawing> ExDrawing::s_instance;

TEST_F(ExDrawing, VariousShapes)
{
    s_instance->VariousShapes();
}

TEST_F(ExDrawing, TypeOfImage)
{
    s_instance->TypeOfImage();
}

TEST_F(ExDrawing, SaveAllImages)
{
    s_instance->SaveAllImages();
}

TEST_F(ExDrawing, ImportImage)
{
    s_instance->ImportImage();
}

TEST_F(ExDrawing, StrokePattern)
{
    s_instance->StrokePattern();
}

TEST_F(ExDrawing, GroupOfShapes)
{
    s_instance->GroupOfShapes();
}

TEST_F(ExDrawing, TextBox)
{
    s_instance->TextBox_();
}

TEST_F(ExDrawing, GetDataFromImage)
{
    s_instance->GetDataFromImage();
}

TEST_F(ExDrawing, ImageData)
{
    s_instance->ImageData_();
}

TEST_F(ExDrawing, ImageSize)
{
    s_instance->ImageSize_();
}

}} // namespace ApiExamples::gtest_test
