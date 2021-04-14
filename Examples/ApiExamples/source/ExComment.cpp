// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExComment.h"

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExComment : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExComment> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExComment>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExComment> ExComment::s_instance;

TEST_F(ExComment, AddCommentWithReply)
{
    s_instance->AddCommentWithReply();
}

TEST_F(ExComment, PrintAllComments)
{
    s_instance->PrintAllComments();
}

TEST_F(ExComment, RemoveCommentReplies)
{
    s_instance->RemoveCommentReplies();
}

TEST_F(ExComment, Done)
{
    s_instance->Done();
}

TEST_F(ExComment, CreateCommentsAndPrintAllInfo)
{
    s_instance->CreateCommentsAndPrintAllInfo();
}

}} // namespace ApiExamples::gtest_test
