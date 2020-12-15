#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Finds all hyperlinks in a Word document and changes their URL and display name.

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeList.h>
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
/// This "facade" class makes it easier to work with a hyperlink field in a Word document.
///
/// A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words
/// consists of several nodes and it might be difficult to work with all those nodes directly.
/// Note this is a simple implementation and will work only if the hyperlink code and name
/// each consist of one Run only.
///
/// [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
///
/// The field code contains a String in one of these formats:
/// HYPERLINK "url"
/// HYPERLINK \l "bookmark name"
///
/// The field result contains text that is displayed to the user.
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
        // Hyperlink display name is stored in the field result which is a Run
        // node between field separator and field end
        auto fieldResult = System::DynamicCast<Run>(mFieldSeparator->get_NextSibling());
        fieldResult->set_Text(value);

        // But sometimes the field result can consist of more than one run, delete these runs
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
        // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        System::Details::ThisProtector __local_self_ref(this);
        //---------------------------------------------------------Self Reference

        if (fieldStart == nullptr)
        {
            throw System::ArgumentNullException(u"fieldStart");
        }
        if (fieldStart->get_FieldType() != FieldType::FieldHyperlink)
        {
            throw System::ArgumentException(u"Field start type must be FieldHyperlink.");
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

        // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs
        String fieldCode = GetTextSameParent(mFieldStart->get_NextSibling(), mFieldSeparator);
        SharedPtr<System::Text::RegularExpressions::Match> match = gRegex->Match(fieldCode.Trim());
        mIsLocal = match->get_Groups()->idx_get(1)->get_Length() > 0;
        // The link is local if \l is present in the field code
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
        // Field code is stored in a Run node between field start and field separator
        auto fieldCode = System::DynamicCast<Run>(mFieldStart->get_NextSibling());
        fieldCode->set_Text(String::Format(u"HYPERLINK {0}\"{1}\"", ((mIsLocal) ? String(u"\\l ") : String(u"")), mTarget));

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
/// Shows how to replace hyperlinks in a Word document.
/// </summary>
class ExReplaceHyperlinks : public ApiExampleBase
{
    //ExSkip

public:
    /// <summary>
    /// Finds all hyperlinks in a Word document and changes their URL and display name.
    /// </summary>
    void Fields()
    {
        // Specify your document name here
        auto doc = MakeObject<Document>(MyDir + u"Hyperlinks.docx");

        // Hyperlinks in a Word documents are fields, select all field start nodes so we can find the hyperlinks
        SharedPtr<NodeList> fieldStarts = doc->SelectNodes(u"//FieldStart");
        for (auto fieldStart : System::IterateOver(fieldStarts->LINQ_OfType<SharedPtr<FieldStart>>()))
        {
            if (fieldStart->get_FieldType() == FieldType::FieldHyperlink)
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

protected:
    static const String NewUrl;
    static const String NewName;
};

} // namespace ApiExamples
//ExEnd
