#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/scope_guard.h>
#include <system/text/regularexpressions/group.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace ApiExamples {

/// <summary>
/// Represents a facade object for a merge field in a Microsoft Word document.
/// </summary>
class MergeField : public System::Object
{
public:
    /// <summary>
    /// Gets the name of the merge field.
    /// </summary>
    String get_Name()
    {
        return GetTextSameParent(mFieldSeparator->get_NextSibling(), mFieldEnd).Trim(MakeArray<char16_t>({u'«', u'»'}));
    }

    /// <summary>
    /// Sets the name of the merge field.
    /// </summary>
    void set_Name(String value)
    {
        // Merge field name is stored in the field result which is a Run
        // node between field separator and field end
        auto fieldResult = System::DynamicCast<Run>(mFieldSeparator->get_NextSibling());
        fieldResult->set_Text(String::Format(u"«{0}»", value));

        // But sometimes the field result can consist of more than one run, delete these runs
        RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);

        UpdateFieldCode(value);
    }

    MergeField(SharedPtr<FieldStart> fieldStart)
    {
        // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        System::Details::ThisProtector __local_self_ref(this);
        //---------------------------------------------------------Self Reference

        if (fieldStart->get_FieldType() != FieldType::FieldMergeField)
        {
            throw System::ArgumentException(u"Field start type must be FieldMergeField.");
        }

        mFieldStart = fieldStart;

        // Find the field separator node
        mFieldSeparator = FindNextSibling(mFieldStart, NodeType::FieldSeparator);
        if (mFieldSeparator == nullptr)
        {
            throw System::InvalidOperationException(u"Cannot find field separator.");
        }

        // Find the field end node. Normally field end will always be found, but in the example document
        // there happens to be a paragraph break included in the hyperlink and this puts the field end
        // in the next paragraph. It will be much more complicated to handle fields which span several
        // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes
        mFieldEnd = FindNextSibling(mFieldSeparator, NodeType::FieldEnd);
    }

private:
    SharedPtr<Node> mFieldStart;
    SharedPtr<Node> mFieldSeparator;
    SharedPtr<Node> mFieldEnd;
    static SharedPtr<System::Text::RegularExpressions::Regex> gRegex;

    void UpdateFieldCode(String fieldName)
    {
        // Field code is stored in a Run node between field start and field separator
        auto fieldCode = System::DynamicCast<Run>(mFieldStart->get_NextSibling());
        SharedPtr<System::Text::RegularExpressions::Match> match = gRegex->Match(fieldCode->get_Text());

        String newFieldCode = String::Format(u" {0}{1} ", match->get_Groups()->idx_get(u"start")->get_Value(), fieldName);
        fieldCode->set_Text(newFieldCode);

        // But sometimes the field code can consist of more than one run, delete these runs
        RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
    }

    /// <summary>
    /// Goes through siblings starting from the start node until it finds a node of the specified type or null.
    /// </summary>
    static SharedPtr<Node> FindNextSibling(SharedPtr<Node> startNode, NodeType nodeType)
    {
        for (SharedPtr<Node> node = startNode; node != nullptr; node = node->get_NextSibling())
        {
            if (node->get_NodeType() == nodeType)
            {
                return node;
            }
        }

        return nullptr;
    }

    /// <summary>
    /// Retrieves text from start up to but not including the end node.
    /// </summary>
    static String GetTextSameParent(SharedPtr<Node> startNode, SharedPtr<Node> endNode)
    {
        if (endNode != nullptr && startNode->get_ParentNode() != endNode->get_ParentNode())
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

    /// <summary>
    /// Removes nodes from start up to but not including the end node.
    /// Start and end are assumed to have the same parent.
    /// </summary>
    static void RemoveSameParent(SharedPtr<Node> startNode, SharedPtr<Node> endNode)
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
};

/// <summary>
/// Shows how to rename merge fields in a Word document.
/// </summary>
class ExRenameMergeFields : public ApiExampleBase
{
    //ExSkip

public:
    /// <summary>
    /// Finds all merge fields in a Word document and changes their names.
    /// </summary>
    void Rename()
    {
        // Create a blank document and insert MERGEFIELDs
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Dear ");
        builder->InsertField(u"MERGEFIELD  FirstName ");
        builder->Write(u" ");
        builder->InsertField(u"MERGEFIELD  LastName ");
        builder->Writeln(u",");
        builder->InsertField(u"MERGEFIELD  CustomGreeting ");

        // Select all field start nodes so we can find the MERGEFIELDs
        SharedPtr<NodeCollection> fieldStarts = doc->GetChildNodes(NodeType::FieldStart, true);
        for (auto fieldStart : System::IterateOver(fieldStarts->LINQ_OfType<SharedPtr<FieldStart>>()))
        {
            if (fieldStart->get_FieldType() == FieldType::FieldMergeField)
            {
                auto mergeField = MakeObject<MergeField>(fieldStart);
                mergeField->set_Name(mergeField->get_Name() + u"_Renamed");
            }
        }

        doc->Save(ArtifactsDir + u"RenameMergeFields.Rename.docx");
    }
};

} // namespace ApiExamples
