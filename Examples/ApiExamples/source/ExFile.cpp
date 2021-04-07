// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExFile.h"

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
namespace ApiExamples { namespace gtest_test {

class ExFile : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExFile> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExFile>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExFile> ExFile::s_instance;

TEST_F(ExFile, CatchFileCorruptedException)
{
    s_instance->CatchFileCorruptedException();
}

TEST_F(ExFile, DetectEncoding)
{
    s_instance->DetectEncoding();
}

TEST_F(ExFile, FileFormatToString)
{
    s_instance->FileFormatToString();
}

TEST_F(ExFile, DetectDocumentEncryption)
{
    s_instance->DetectDocumentEncryption();
}

TEST_F(ExFile, DetectDigitalSignatures)
{
    s_instance->DetectDigitalSignatures();
}

TEST_F(ExFile, SaveToDetectedFileFormat)
{
    s_instance->SaveToDetectedFileFormat();
}

TEST_F(ExFile, DetectFileFormat_SaveFormatToLoadFormat)
{
    s_instance->DetectFileFormat_SaveFormatToLoadFormat();
}

TEST_F(ExFile, ExtractImages)
{
    s_instance->ExtractImages();
}

}} // namespace ApiExamples::gtest_test
