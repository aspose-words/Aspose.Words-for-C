// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentProperties.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Properties;
namespace ApiExamples { namespace gtest_test {

class ExDocumentProperties : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDocumentProperties> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDocumentProperties>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDocumentProperties> ExDocumentProperties::s_instance;

TEST_F(ExDocumentProperties, BuiltIn)
{
    s_instance->BuiltIn();
}

TEST_F(ExDocumentProperties, Custom)
{
    s_instance->Custom();
}

TEST_F(ExDocumentProperties, Description)
{
    s_instance->Description();
}

TEST_F(ExDocumentProperties, Origin)
{
    s_instance->Origin();
}

TEST_F(ExDocumentProperties, Content)
{
    s_instance->Content();
}

TEST_F(ExDocumentProperties, Thumbnail)
{
    s_instance->Thumbnail();
}

TEST_F(ExDocumentProperties, HyperlinkBase)
{
    s_instance->HyperlinkBase();
}

TEST_F(ExDocumentProperties, HeadingPairs)
{
    s_instance->HeadingPairs();
}

TEST_F(ExDocumentProperties, Security)
{
    s_instance->Security();
}

TEST_F(ExDocumentProperties, CustomNamedAccess)
{
    s_instance->CustomNamedAccess();
}

TEST_F(ExDocumentProperties, LinkCustomDocumentPropertiesToBookmark)
{
    s_instance->LinkCustomDocumentPropertiesToBookmark();
}

TEST_F(ExDocumentProperties, DocumentPropertyCollection)
{
    s_instance->DocumentPropertyCollection();
}

TEST_F(ExDocumentProperties, PropertyTypes)
{
    s_instance->PropertyTypes();
}

}} // namespace ApiExamples::gtest_test
