// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRenameMergeFields.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace ApiExamples {

System::SharedPtr<System::Text::RegularExpressions::Regex> MergeField::gRegex =
    System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+");

namespace gtest_test {

class ExRenameMergeFields : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExRenameMergeFields> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExRenameMergeFields>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExRenameMergeFields> ExRenameMergeFields::s_instance;

TEST_F(ExRenameMergeFields, Rename)
{
    s_instance->Rename();
}

} // namespace gtest_test

} // namespace ApiExamples
