// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExProperties.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <system/shared_ptr.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
namespace ApiExamples { namespace gtest_test {

class ExProperties : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExProperties> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExProperties>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExProperties> ExProperties::s_instance;

TEST_F(ExProperties, BuiltIn)
{
    s_instance->BuiltIn();
}

TEST_F(ExProperties, Custom)
{
    s_instance->Custom();
}

TEST_F(ExProperties, Description)
{
    s_instance->Description();
}

TEST_F(ExProperties, Origin)
{
    s_instance->Origin();
}

TEST_F(ExProperties, Content)
{
    s_instance->Content();
}

TEST_F(ExProperties, Thumbnail)
{
    s_instance->Thumbnail();
}

TEST_F(ExProperties, HyperlinkBase)
{
    s_instance->HyperlinkBase();
}

TEST_F(ExProperties, HeadingPairs)
{
    s_instance->HeadingPairs();
}

TEST_F(ExProperties, Security)
{
    s_instance->Security();
}

TEST_F(ExProperties, CustomNamedAccess)
{
    s_instance->CustomNamedAccess();
}

TEST_F(ExProperties, LinkCustomDocumentPropertiesToBookmark)
{
    s_instance->LinkCustomDocumentPropertiesToBookmark();
}

TEST_F(ExProperties, DocumentPropertyCollection)
{
    s_instance->DocumentPropertyCollection();
}

TEST_F(ExProperties, PropertyTypes)
{
    s_instance->PropertyTypes();
}

}} // namespace ApiExamples::gtest_test
