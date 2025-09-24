// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExRenameMergeFields.h"

#include <system/text/string_builder.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/group.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Fields;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2701642915u, ::Aspose::Words::ApiExamples::ExRenameMergeFields, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExRenameMergeFields : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExRenameMergeFields> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExRenameMergeFields>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExRenameMergeFields> ExRenameMergeFields::s_instance;

} // namespace gtest_test

void ExRenameMergeFields::Rename()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Dear ");
    builder->InsertField(u"MERGEFIELD  FirstName ");
    builder->Write(u" ");
    builder->InsertField(u"MERGEFIELD  LastName ");
    builder->Writeln(u",");
    builder->InsertField(u"MERGEFIELD  CustomGreeting ");
    
    // Select all field start nodes so we can find the MERGEFIELDs.
    System::SharedPtr<Aspose::Words::NodeCollection> fieldStarts = doc->GetChildNodes(Aspose::Words::NodeType::FieldStart, true);
    for (auto&& fieldStart : System::IterateOver(fieldStarts->LINQ_OfType<System::SharedPtr<Aspose::Words::Fields::FieldStart> >()))
    {
        if (fieldStart->get_FieldType() == Aspose::Words::Fields::FieldType::FieldMergeField)
        {
            auto mergeField = System::MakeObject<Aspose::Words::ApiExamples::MergeField>(fieldStart);
            mergeField->set_Name(mergeField->get_Name() + u"_Renamed");
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"RenameMergeFields.Rename.docx");
}

namespace gtest_test
{

TEST_F(ExRenameMergeFields, Rename)
{
    s_instance->Rename();
}

} // namespace gtest_test

RTTI_INFO_IMPL_HASH(1854326943u, ::Aspose::Words::ApiExamples::MergeField, ThisTypeBaseTypesInfo);

System::String MergeField::get_Name()
{
    return GetTextSameParent(mFieldSeparator->get_NextSibling(), mFieldEnd).Trim(System::MakeArray<char16_t>({u'«', u'»'}));
}

void MergeField::set_Name(System::String value)
{
    // Merge field name is stored in the field result which is a Run 
    // node between field separator and field end.
    auto fieldResult = System::ExplicitCast<Aspose::Words::Run>(mFieldSeparator->get_NextSibling());
    fieldResult->set_Text(System::String::Format(u"«{0}»", value));
    
    // But sometimes the field result can consist of more than one run, delete these runs.
    RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);
    
    UpdateFieldCode(value);
}

System::SharedPtr<System::Text::RegularExpressions::Regex>& MergeField::gRegex()
{
    static System::SharedPtr<System::Text::RegularExpressions::Regex> value = System::MakeObject<System::Text::RegularExpressions::Regex>(u"\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+");
    return value;
}

MergeField::MergeField(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart)
{
    if (fieldStart->get_FieldType() != Aspose::Words::Fields::FieldType::FieldMergeField)
    {
        throw System::ArgumentException(u"Field start type must be FieldMergeField.");
    }
    
    mFieldStart = fieldStart;
    
    // Find the field separator node.
    mFieldSeparator = FindNextSibling(mFieldStart, Aspose::Words::NodeType::FieldSeparator);
    if (mFieldSeparator == nullptr)
    {
        throw System::InvalidOperationException(u"Cannot find field separator.");
    }
    
    // Find the field end node. Normally field end will always be found, but in the example document 
    // there happens to be a paragraph break included in the hyperlink and this puts the field end 
    // in the next paragraph. It will be much more complicated to handle fields which span several 
    // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
    mFieldEnd = FindNextSibling(mFieldSeparator, Aspose::Words::NodeType::FieldEnd);
}

void MergeField::UpdateFieldCode(System::String fieldName)
{
    // Field code is stored in a Run node between field start and field separator.
    auto fieldCode = System::ExplicitCast<Aspose::Words::Run>(mFieldStart->get_NextSibling());
    System::SharedPtr<System::Text::RegularExpressions::Match> match = gRegex()->Match(fieldCode->get_Text());
    
    System::String newFieldCode = System::String::Format(u" {0}{1} ", match->get_Groups()->idx_get(u"start")->get_Value(), fieldName);
    fieldCode->set_Text(newFieldCode);
    
    // But sometimes the field code can consist of more than one run, delete these runs.
    RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
}

System::SharedPtr<Aspose::Words::Node> MergeField::FindNextSibling(System::SharedPtr<Aspose::Words::Node> startNode, Aspose::Words::NodeType nodeType)
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

System::String MergeField::GetTextSameParent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode)
{
    if (endNode != nullptr && startNode->get_ParentNode() != endNode->get_ParentNode())
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

void MergeField::RemoveSameParent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode)
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
