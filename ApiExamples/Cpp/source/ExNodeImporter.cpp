// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExNodeImporter.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples { namespace gtest_test {

class ExNodeImporter : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExNodeImporter> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExNodeImporter>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExNodeImporter> ExNodeImporter::s_instance;

TEST_F(ExNodeImporter, KeepSourceNumbering)
{
    s_instance->KeepSourceNumbering();
}

TEST_F(ExNodeImporter, InsertAtBookmark)
{
    s_instance->InsertAtBookmark();
}

TEST_F(ExNodeImporter, InsertAtMailMerge)
{
    s_instance->InsertAtMailMerge();
}

}} // namespace ApiExamples::gtest_test
