// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Finds all hyperlinks in a Word document and changes their URL and display name.
#include "ExReplaceHyperlinks.h"

#include <system/text/string_builder.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/group.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/scope_guard.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeList.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

namespace ApiExamples { class Hyperlink; }

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
namespace ApiExamples {

const String ExReplaceHyperlinks::NewUrl = u"http://www.aspose.com";
const String ExReplaceHyperlinks::NewName = u"Aspose - The .NET & Java Component Publisher";

namespace gtest_test
{

class ExReplaceHyperlinks : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExReplaceHyperlinks> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExReplaceHyperlinks>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExReplaceHyperlinks> ExReplaceHyperlinks::s_instance;

} // namespace gtest_test

void ExReplaceHyperlinks::Fields()
{
    // Specify your document name here
    auto doc = MakeObject<Document>(MyDir + u"Hyperlinks.docx");

    // Hyperlinks in a Word documents are fields, select all field start nodes so we can find the hyperlinks
    SharedPtr<NodeList> fieldStarts = doc->SelectNodes(u"//FieldStart");
    for (auto fieldStart : System::IterateOver(fieldStarts->LINQ_OfType<SharedPtr<FieldStart> >()))
    {
        if (System::ObjectExt::Equals(fieldStart->get_FieldType(), Aspose::Words::Fields::FieldType::FieldHyperlink))
        {
            // The field is a hyperlink field, use the "facade" class to help to deal with the field
            auto hyperlink = MakeObject<Hyperlink>(fieldStart);

            // Some hyperlinks can be local (links to bookmarks inside the document), ignore these
            if (hyperlink->get_IsLocal())
            {
                continue;
            }

            // The Hyperlink class allows to set the target URL and the display name
            // of the link easily by setting the properties
            hyperlink->set_Target(NewUrl);
            hyperlink->set_Name(NewName);
        }
    }

    doc->Save(ArtifactsDir + u"ReplaceHyperlinks.Fields.docx");
}

namespace gtest_test
{

TEST_F(ExReplaceHyperlinks, Fields)
{
    s_instance->Fields();
}

} // namespace gtest_test

RTTI_INFO_IMPL_HASH(106621379u, ::ApiExamples::Hyperlink, ThisTypeBaseTypesInfo);

SharedPtr<System::Text::RegularExpressions::Regex> Hyperlink::gRegex = MakeObject<System::Text::RegularExpressions::Regex>(String(u"\\S+") + u"\\s+" + u"(?:\"\"\\s+)?" + u"(\\\\l\\s+)?" + u"\"" + u"([^\"]+)" + u"\"");

String Hyperlink::get_Name()
{
    return GetTextSameParent(mFieldSeparator, mFieldEnd);
}

void Hyperlink::set_Name(String value)
{
    // Hyperlink display name is stored in the field result which is a Run
    // node between field separator and field end
    auto fieldResult = System::DynamicCast<Aspose::Words::Run>(mFieldSeparator->get_NextSibling());
    fieldResult->set_Text(value);

    // But sometimes the field result can consist of more than one run, delete these runs
    RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);
}

String Hyperlink::get_Target()
{
    return mTarget;
}

void Hyperlink::set_Target(String value)
{
    mTarget = value;
    UpdateFieldCode();
}

bool Hyperlink::get_IsLocal()
{
    return mIsLocal;
}

void Hyperlink::set_IsLocal(bool value)
{
    mIsLocal = value;
    UpdateFieldCode();
}

Hyperlink::Hyperlink(SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart) : mIsLocal(false)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    if (fieldStart == nullptr)
    {
        throw System::ArgumentNullException(u"fieldStart");
    }
    if (!System::ObjectExt::Equals(fieldStart->get_FieldType(), Aspose::Words::Fields::FieldType::FieldHyperlink))
    {
        throw System::ArgumentException(u"Field start type must be FieldHyperlink.");
    }

    mFieldStart = fieldStart;

    // Find the field separator node
    mFieldSeparator = FindNextSibling(mFieldStart, Aspose::Words::NodeType::FieldSeparator);
    if (mFieldSeparator == nullptr)
    {
        throw System::InvalidOperationException(u"Cannot find field separator.");
    }

    // Find the field end node. Normally field end will always be found, but in the example document
    // there happens to be a paragraph break included in the hyperlink and this puts the field end
    // in the next paragraph. It will be much more complicated to handle fields which span several
    // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes
    mFieldEnd = FindNextSibling(mFieldSeparator, Aspose::Words::NodeType::FieldEnd);

    // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs
    String fieldCode = GetTextSameParent(mFieldStart->get_NextSibling(), mFieldSeparator);
    SharedPtr<System::Text::RegularExpressions::Match> match = gRegex->Match(fieldCode.Trim());
    mIsLocal = match->get_Groups()->idx_get(1)->get_Length() > 0;
    //The link is local if \l is present in the field code
    mTarget = match->get_Groups()->idx_get(2)->get_Value();
}

void Hyperlink::UpdateFieldCode()
{
    // Field code is stored in a Run node between field start and field separator
    auto fieldCode = System::DynamicCast<Aspose::Words::Run>(mFieldStart->get_NextSibling());
    fieldCode->set_Text(String::Format(u"HYPERLINK {0}\"{1}\"",((mIsLocal) ? String(u"\\l ") : String(u"")),mTarget));

    // But sometimes the field code can consist of more than one run, delete these runs
    RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
}

SharedPtr<Aspose::Words::Node> Hyperlink::FindNextSibling(SharedPtr<Aspose::Words::Node> startNode, Aspose::Words::NodeType nodeType)
{
    for (SharedPtr<Node> node = startNode; node != nullptr; node = node->get_NextSibling())
    {
        if (System::ObjectExt::Equals(node->get_NodeType(), nodeType))
        {
            return node;
        }
    }

    return nullptr;
}

String Hyperlink::GetTextSameParent(SharedPtr<Aspose::Words::Node> startNode, SharedPtr<Aspose::Words::Node> endNode)
{
    if ((endNode != nullptr) && (startNode->get_ParentNode() != endNode->get_ParentNode()))
    {
        throw System::ArgumentException(u"Start and end nodes are expected to have the same parent.");
    }

    auto builder = MakeObject<System::Text::StringBuilder>();
    for (SharedPtr<Node> child = startNode; !System::ObjectExt::Equals(child, endNode); child = child->get_NextSibling())
    {
        builder->Append(child->GetText());
    }

    return builder->ToString();
}

void Hyperlink::RemoveSameParent(SharedPtr<Aspose::Words::Node> startNode, SharedPtr<Aspose::Words::Node> endNode)
{
    if (endNode != nullptr && startNode->get_ParentNode() != endNode->get_ParentNode())
    {
        throw System::ArgumentException(u"Start and end nodes are expected to have the same parent.");
    }

    SharedPtr<Node> curChild = startNode;
    while ((curChild != nullptr) && (curChild != endNode))
    {
        SharedPtr<Node> nextChild = curChild->get_NextSibling();
        curChild->Remove();
        curChild = nextChild;
    }
}

System::Object::shared_members_type ApiExamples::Hyperlink::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::Hyperlink::mFieldStart", this->mFieldStart);
    result.Add("ApiExamples::Hyperlink::mFieldSeparator", this->mFieldSeparator);
    result.Add("ApiExamples::Hyperlink::mFieldEnd", this->mFieldEnd);

    return result;
}

} // namespace ApiExamples
