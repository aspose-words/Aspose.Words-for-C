// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Shows how to find all hyperlinks in a Word document, and then change their URLs and display names.
#include "ExReplaceHyperlinks.h"

#include <system/text/string_builder.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/group.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeList.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Fields;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3705082897u, ::Aspose::Words::ApiExamples::ExReplaceHyperlinks, ThisTypeBaseTypesInfo);

const System::String& ExReplaceHyperlinks::NewUrl()
{
    static const System::String value = u"http://www.aspose.com";
    return value;
}

const System::String& ExReplaceHyperlinks::NewName()
{
    static const System::String value = u"Aspose - The .NET & Java Component Publisher";
    return value;
}


namespace gtest_test
{

class ExReplaceHyperlinks : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExReplaceHyperlinks> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExReplaceHyperlinks>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExReplaceHyperlinks> ExReplaceHyperlinks::s_instance;

} // namespace gtest_test

void ExReplaceHyperlinks::Fields()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Hyperlinks.docx");
    
    // Hyperlinks in a Word documents are fields. To begin looking for hyperlinks, we must first find all the fields.
    // Use the "SelectNodes" method to find all the fields in the document via an XPath.
    System::SharedPtr<Aspose::Words::NodeList> fieldStarts = doc->SelectNodes(u"//FieldStart");
    
    for (auto&& fieldStart : System::IterateOver(fieldStarts->LINQ_OfType<System::SharedPtr<Aspose::Words::Fields::FieldStart> >()))
    {
        if (fieldStart->get_FieldType() == Aspose::Words::Fields::FieldType::FieldHyperlink)
        {
            auto hyperlink = System::MakeObject<Aspose::Words::ApiExamples::Hyperlink>(fieldStart);
            
            // Hyperlinks that link to bookmarks do not have URLs.
            if (hyperlink->get_IsLocal())
            {
                continue;
            }
            
            // Give each URL hyperlink a new URL and name.
            hyperlink->set_Target(NewUrl());
            hyperlink->set_Name(NewName());
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"ReplaceHyperlinks.Fields.docx");
}

namespace gtest_test
{

TEST_F(ExReplaceHyperlinks, Fields)
{
    s_instance->Fields();
}

} // namespace gtest_test

RTTI_INFO_IMPL_HASH(2447104009u, ::Aspose::Words::ApiExamples::Hyperlink, ThisTypeBaseTypesInfo);

System::String Hyperlink::get_Name()
{
    return GetTextSameParent(mFieldSeparator, mFieldEnd);
}

void Hyperlink::set_Name(System::String value)
{
    // Hyperlink display name is stored in the field result, which is a Run 
    // node between field separator and field end.
    auto fieldResult = System::ExplicitCast<Aspose::Words::Run>(mFieldSeparator->get_NextSibling());
    fieldResult->set_Text(value);
    
    // If the field result consists of more than one run, delete these runs.
    RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);
}

System::String Hyperlink::get_Target() const
{
    return mTarget;
}

void Hyperlink::set_Target(System::String value)
{
    mTarget = value;
    UpdateFieldCode();
}

bool Hyperlink::get_IsLocal() const
{
    return mIsLocal;
}

void Hyperlink::set_IsLocal(bool value)
{
    mIsLocal = value;
    UpdateFieldCode();
}

System::SharedPtr<System::Text::RegularExpressions::Regex>& Hyperlink::gRegex()
{
    static System::SharedPtr<System::Text::RegularExpressions::Regex> value = System::MakeObject<System::Text::RegularExpressions::Regex>(System::String(u"\\S+") + u"\\s+" + u"(?:\"\"\\s+)?" + u"(\\\\l\\s+)?" + u"\"" + u"([^\"]+)" + u"\"");
    return value;
}

Hyperlink::Hyperlink(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart) : mIsLocal(false)
{
    if (fieldStart == nullptr)
    {
        throw System::ArgumentNullException(u"fieldStart");
    }
    if (fieldStart->get_FieldType() != Aspose::Words::Fields::FieldType::FieldHyperlink)
    {
        throw System::ArgumentException(u"Field start type must be FieldHyperlink.");
    }
    
    mFieldStart = fieldStart;
    
    // Find the field separator node.
    mFieldSeparator = FindNextSibling(mFieldStart, Aspose::Words::NodeType::FieldSeparator);
    if (mFieldSeparator == nullptr)
    {
        throw System::InvalidOperationException(u"Cannot find field separator.");
    }
    
    // Normally, we can always find the field's end node, but the example document 
    // contains a paragraph break inside a hyperlink, which puts the field end 
    // in the next paragraph. It will be much more complicated to handle fields which span several 
    // paragraphs correctly. In this case allowing field end to be null is enough.
    mFieldEnd = FindNextSibling(mFieldSeparator, Aspose::Words::NodeType::FieldEnd);
    
    // Field code looks something like "HYPERLINK "http:\\www.myurl.com"", but it can consist of several runs.
    System::String fieldCode = GetTextSameParent(mFieldStart->get_NextSibling(), mFieldSeparator);
    System::SharedPtr<System::Text::RegularExpressions::Match> match = gRegex()->Match(fieldCode.Trim());
    
    // The hyperlink is local if \l is present in the field code.
    mIsLocal = match->get_Groups()->idx_get(1)->get_Length() > 0;
    mTarget = match->get_Groups()->idx_get(2)->get_Value();
}

void Hyperlink::UpdateFieldCode()
{
    // A field's field code is in a Run node between the field's start node and field separator.
    auto fieldCode = System::ExplicitCast<Aspose::Words::Run>(mFieldStart->get_NextSibling());
    fieldCode->set_Text(System::String::Format(u"HYPERLINK {0}\"{1}\"", ((mIsLocal) ? System::String(u"\\l ") : System::String(u"")), mTarget));
    
    // If the field code consists of more than one run, delete these runs.
    RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
}

System::SharedPtr<Aspose::Words::Node> Hyperlink::FindNextSibling(System::SharedPtr<Aspose::Words::Node> startNode, Aspose::Words::NodeType nodeType)
{
    for (System::SharedPtr<Aspose::Words::Node> node = startNode; node != nullptr; node = node->get_NextSibling())
    {
        if (node->get_NodeType() == nodeType)
        {
            return node;
        }
    }
    
    return nullptr;
}

System::String Hyperlink::GetTextSameParent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode)
{
    if ((endNode != nullptr) && (startNode->get_ParentNode() != endNode->get_ParentNode()))
    {
        throw System::ArgumentException(u"Start and end nodes are expected to have the same parent.");
    }
    
    System::SharedPtr<System::Text::StringBuilder> builder = System::MakeObject<System::Text::StringBuilder>();
    for (System::SharedPtr<Aspose::Words::Node> child = startNode; !System::ObjectExt::Equals(child, endNode); child = child->get_NextSibling())
    {
        builder->Append(child->GetText());
    }
    
    return builder->ToString();
}

void Hyperlink::RemoveSameParent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode)
{
    if (endNode != nullptr && startNode->get_ParentNode() != endNode->get_ParentNode())
    {
        throw System::ArgumentException(u"Start and end nodes are expected to have the same parent.");
    }
    
    System::SharedPtr<Aspose::Words::Node> curChild = startNode;
    while ((curChild != nullptr) && (curChild != endNode))
    {
        System::SharedPtr<Aspose::Words::Node> nextChild = curChild->get_NextSibling();
        curChild->Remove();
        curChild = nextChild;
    }
}

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
