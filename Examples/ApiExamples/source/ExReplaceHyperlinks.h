#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Shows how to find all hyperlinks in a Word document, and then change their URLs and display names.

#include <system/text/regularexpressions/regex.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fields;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExReplaceHyperlinks : public ApiExampleBase
{
    typedef ExReplaceHyperlinks ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void Fields();
    
protected:

    static const System::String& NewUrl();
    static const System::String& NewName();
    
};

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
    typedef Hyperlink ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Gets or sets the display name of the hyperlink.
    /// </summary>
    System::String get_Name();
    /// <summary>
    /// Gets or sets the display name of the hyperlink.
    /// </summary>
    void set_Name(System::String value);
    /// <summary>
    /// Gets or sets the target URL or bookmark name of the hyperlink.
    /// </summary>
    System::String get_Target() const;
    /// <summary>
    /// Gets or sets the target URL or bookmark name of the hyperlink.
    /// </summary>
    void set_Target(System::String value);
    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL.
    /// </summary>
    bool get_IsLocal() const;
    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL.
    /// </summary>
    void set_IsLocal(bool value);
    
    Hyperlink(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart);
    
private:

    System::SharedPtr<Aspose::Words::Node> mFieldStart;
    System::SharedPtr<Aspose::Words::Node> mFieldSeparator;
    System::SharedPtr<Aspose::Words::Node> mFieldEnd;
    bool mIsLocal;
    System::String mTarget;
    
    static System::SharedPtr<System::Text::RegularExpressions::Regex>& gRegex();
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
    /// Assumes that the start and end nodes have the same parent.
    /// </summary>
    static void RemoveSameParent(System::SharedPtr<Aspose::Words::Node> startNode, System::SharedPtr<Aspose::Words::Node> endNode);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
//ExEnd


