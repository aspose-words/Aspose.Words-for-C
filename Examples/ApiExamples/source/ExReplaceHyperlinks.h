#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Shows how to find all hyperlinks in a Word document, and then change their URLs and display names.

#include <cstdint>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeList.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/text/regularexpressions/group.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/string_builder.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace ApiExamples {

/// <summary>
/// HYPERLINK fields contain and display hyperlinks in the document body. A field in Aspose.Words
/// consists of several nodes, and it might be difficult to work with all those nodes directly.
/// This implementation will work only if the hyperlink code and name each consist of only one Run node.
///
/// The node structure for fields is as follows:
///
/// [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
///
/// Below are two example field codes of HYPERLINK fields:
/// HYPERLINK "url"
/// HYPERLINK \l "bookmark name"
///
/// A field's "Result" property contains text that the field displays in the document body to the user.
/// </summary>
class Hyperlink : public System::Object
{
public:
    /// <summary>
    /// Gets the display name of the hyperlink.
    /// </summary>
    String get_Name()
    {
        return GetTextSameParent(mFieldSeparator, mFieldEnd);
    }

    /// <summary>
    /// Sets the display name of the hyperlink.
    /// </summary>
    void set_Name(String value)
    {
        // Hyperlink display name is stored in the field result, which is a Run
        // node between field separator and field end.
        auto fieldResult = System::DynamicCast<Run>(mFieldSeparator->get_NextSibling());
        fieldResult->set_Text(value);

        // If the field result consists of more than one run, delete these runs.
        RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);
    }

    /// <summary>
    /// Gets the target URL or bookmark name of the hyperlink.
    /// </summary>
    String get_Target()
    {
        return mTarget;
    }

    /// <summary>
    /// Sets the target URL or bookmark name of the hyperlink.
    /// </summary>
    void set_Target(String value)
    {
        mTarget = value;
        UpdateFieldCode();
    }

    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL.
    /// </summary>
    bool get_IsLocal()
    {
        return mIsLocal;
    }

    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL.
    /// </summary>
    void set_IsLocal(bool value)
    {
        mIsLocal = value;
        UpdateFieldCode();
    }

    Hyperlink(SharedPtr<FieldStart> fieldStart) : mIsLocal(false)
    {
        if (fieldStart == nullptr)
        {
            throw System::ArgumentNullException(u"fieldStart");
        }
        if (fieldStart->get_FieldType() != FieldType::FieldHyperlink)
        {
            throw System::ArgumentException(u"Field start type must be FieldHyperlink.");
        }

        mFieldStart = fieldStart;

        // Find the field separator node.
        mFieldSeparator = FindNextSibling(mFieldStart, NodeType::FieldSeparator);
        if (mFieldSeparator == nullptr)
        {
            throw System::InvalidOperationException(u"Cannot find field separator.");
        }

        // Normally, we can always find the field's end node, but the example document
        // contains a paragraph break inside a hyperlink, which puts the field end
        // in the next paragraph. It will be much more complicated to handle fields which span several
        // paragraphs correctly. In this case allowing field end to be null is enough.
        mFieldEnd = FindNextSibling(mFieldSeparator, NodeType::FieldEnd);

        // Field code looks something like "HYPERLINK "http:\\www.myurl.com"", but it can consist of several runs.
        String fieldCode = GetTextSameParent(mFieldStart->get_NextSibling(), mFieldSeparator);
        SharedPtr<System::Text::RegularExpressions::Match> match = gRegex->Match(fieldCode.Trim());

        // The hyperlink is local if \l is present in the field code.
        mIsLocal = match->get_Groups()->idx_get(1)->get_Length() > 0;
        mTarget = match->get_Groups()->idx_get(2)->get_Value();
    }

private:
    SharedPtr<Node> mFieldStart;
    SharedPtr<Node> mFieldSeparator;
    SharedPtr<Node> mFieldEnd;
    bool mIsLocal;
    String mTarget;
    static SharedPtr<System::Text::RegularExpressions::Regex> gRegex;

    void UpdateFieldCode()
    {
        // A field's field code is in a Run node between the field's start node and field separator.
        auto fieldCode = System::DynamicCast<Run>(mFieldStart->get_NextSibling());
        fieldCode->set_Text(String::Format(u"HYPERLINK {0}\"{1}\"", ((mIsLocal) ? String(u"\\l ") : String(u"")), mTarget));

        // If the field code consists of more than one run, delete these runs.
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

    /// <summary>
    /// Removes nodes from start up to but not including the end node.
    /// Assumes that the start and end nodes have the same parent.
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

class ExReplaceHyperlinks : public ApiExampleBase
{
    //ExSkip

public:
    void Fields()
    {
        auto doc = MakeObject<Document>(MyDir + u"Hyperlinks.docx");

        // Hyperlinks in a Word documents are fields. To begin looking for hyperlinks, we must first find all the fields.
        // Use the "SelectNodes" method to find all the fields in the document via an XPath.
        SharedPtr<NodeList> fieldStarts = doc->SelectNodes(u"//FieldStart");

        for (const auto& fieldStart : System::IterateOver(fieldStarts->LINQ_OfType<SharedPtr<FieldStart>>()))
        {
            if (fieldStart->get_FieldType() == FieldType::FieldHyperlink)
            {
                auto hyperlink = MakeObject<Hyperlink>(fieldStart);

                // Hyperlinks that link to bookmarks do not have URLs.
                if (hyperlink->get_IsLocal())
                {
                    continue;
                }

                // Give each URL hyperlink a new URL and name.
                hyperlink->set_Target(NewUrl);
                hyperlink->set_Name(NewName);
            }
        }

        doc->Save(ArtifactsDir + u"ReplaceHyperlinks.Fields.docx");
    }

    static const String NewUrl;
    static const String NewName;
};

} // namespace ApiExamples
//ExEnd
