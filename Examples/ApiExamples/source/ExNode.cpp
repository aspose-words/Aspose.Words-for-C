// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExNode.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExNode : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExNode> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExNode>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExNode> ExNode::s_instance;

TEST_F(ExNode, CloneCompositeNode)
{
    s_instance->CloneCompositeNode();
}

TEST_F(ExNode, GetParentNode)
{
    s_instance->GetParentNode();
}

TEST_F(ExNode, OwnerDocument)
{
    s_instance->OwnerDocument();
}

TEST_F(ExNode, ChildNodesEnumerate)
{
    s_instance->ChildNodesEnumerate();
}

TEST_F(ExNode, RecurseChildren)
{
    s_instance->RecurseChildren();
}

TEST_F(ExNode, RemoveNodes)
{
    s_instance->RemoveNodes();
}

TEST_F(ExNode, EnumNextSibling)
{
    s_instance->EnumNextSibling();
}

TEST_F(ExNode, TypedAccess)
{
    s_instance->TypedAccess();
}

TEST_F(ExNode, RemoveChild)
{
    s_instance->RemoveChild();
}

TEST_F(ExNode, CreateAndAddParagraphNode)
{
    s_instance->CreateAndAddParagraphNode();
}

TEST_F(ExNode, RemoveSmartTagsFromCompositeNode)
{
    s_instance->RemoveSmartTagsFromCompositeNode();
}

TEST_F(ExNode, GetIndexOfNode)
{
    s_instance->GetIndexOfNode();
}

TEST_F(ExNode, ConvertNodeToHtmlWithDefaultOptions)
{
    s_instance->ConvertNodeToHtmlWithDefaultOptions();
}

TEST_F(ExNode, TypedNodeCollectionToArray)
{
    s_instance->TypedNodeCollectionToArray();
}

TEST_F(ExNode, NodeEnumerationHotRemove)
{
    s_instance->NodeEnumerationHotRemove();
}

TEST_F(ExNode, NodeChangingCallback)
{
    s_instance->NodeChangingCallback();
}

TEST_F(ExNode, NodeCollection)
{
    s_instance->NodeCollection_();
}

}} // namespace ApiExamples::gtest_test
