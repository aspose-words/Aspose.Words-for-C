// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported
#include "ExDocumentBase.h"

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Loading;
namespace ApiExamples { namespace gtest_test {

class ExDocumentBase : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDocumentBase> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDocumentBase>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDocumentBase> ExDocumentBase::s_instance;

TEST_F(ExDocumentBase, Constructor)
{
    s_instance->Constructor();
}

TEST_F(ExDocumentBase, SetPageColor)
{
    s_instance->SetPageColor();
}

TEST_F(ExDocumentBase, ImportNode)
{
    s_instance->ImportNode();
}

TEST_F(ExDocumentBase, ImportNodeCustom)
{
    s_instance->ImportNodeCustom();
}

TEST_F(ExDocumentBase, BackgroundShape)
{
    s_instance->BackgroundShape();
}

TEST_F(ExDocumentBase, ResourceLoadingCallback)
{
    s_instance->ResourceLoadingCallback();
}

}} // namespace ApiExamples::gtest_test
