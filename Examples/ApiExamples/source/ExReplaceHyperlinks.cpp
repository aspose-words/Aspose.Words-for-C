// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Shows how to find all hyperlinks in a Word document, and then change their URLs and display names.
#include "ExReplaceHyperlinks.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace ApiExamples {

System::SharedPtr<System::Text::RegularExpressions::Regex> Hyperlink::gRegex = System::MakeObject<System::Text::RegularExpressions::Regex>(
    System::String(u"\\S+") + u"\\s+" + u"(?:\"\"\\s+)?" + u"(\\\\l\\s+)?" + u"\"" + u"([^\"]+)" + u"\"");

const System::String ExReplaceHyperlinks::NewUrl = u"http://www.aspose.com";
const System::String ExReplaceHyperlinks::NewName = u"Aspose - The .NET & Java Component Publisher";

namespace gtest_test {

class ExReplaceHyperlinks : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExReplaceHyperlinks> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExReplaceHyperlinks>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExReplaceHyperlinks> ExReplaceHyperlinks::s_instance;

TEST_F(ExReplaceHyperlinks, Fields)
{
    s_instance->Fields();
}

} // namespace gtest_test

} // namespace ApiExamples
