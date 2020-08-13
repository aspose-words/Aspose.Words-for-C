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

#include <system/text/regularexpressions/regex.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class Node; } }
namespace Aspose { namespace Words { namespace Fields { class FieldStart; } } }
namespace Aspose { namespace Words { enum class NodeType; } }

namespace ApiExamples {

/// <summary>
/// Shows how to replace hyperlinks in a Word document.
/// </summary>
class ExReplaceHyperlinks : public ApiExampleBase
{
public:

    /// <summary>
    /// Finds all hyperlinks in a Word document and changes their URL and display name.
    /// </summary>
    void Fields();
    
protected:

    static const System::String NewUrl;
    static const System::String NewName;
    
};

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
    typedef Hyperlink ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Gets the display name of the hyperlink.
    /// </summary>
    System::String get_Name();
    /// <summary>
    /// Sets the display name of the hyperlink.
    /// </summary>
    void set_Name(System::String value);
    /// <summary>
    /// Gets the target url or bookmark name of the hyperlink.
    /// </summary>
    System::String get_Target();
    /// <summary>
    /// Sets the target url or bookmark name of the hyperlink.
    /// </summary>
    void set_Target(System::String value);
    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a url.
    /// </summary>
    bool get_IsLocal();
    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a url.
    /// </summary>
    void set_IsLocal(bool value);
    
    Hyperlink(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart);
    
protected:

    System::Object::shared_members_type GetSharedMembers() override;
    
private:

    System::SharedPtr<Aspose::Words::Node> mFieldStart;
    System::SharedPtr<Aspose::Words::Node> mFieldSeparator;
    System::SharedPtr<Aspose::Words::Node> mFieldEnd;
    bool mIsLocal;
    System::String mTarget;
    static System::SharedPtr<System::Text::RegularExpressions::Regex> gRegex;
    
    void UpdateFieldCode();
    /// <summary>
    /// Goes through siblings starting from the start node until it finds a node of the specified type or null.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Node> FindNextSibling(System::SharedPtr<Aspose::Words::Node> startNode, Aspose::Words::NodeType nodeType);
    /// <summary>
    /// Retrieves text from start up to but not including the end node.
    /// </summary>
    static System::String GetTextSameParent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode);
    /// <summary>
    /// Removes nodes from start up to but not including the end node.
    /// Start and end are assumed to have the same parent.
    /// </summary>
    static void RemoveSameParent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode);
    
};

} // namespace ApiExamples
//ExEnd


