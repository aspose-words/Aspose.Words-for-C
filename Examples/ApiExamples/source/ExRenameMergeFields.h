#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/regularexpressions/regex.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fields;

namespace Aspose {

namespace Words {

namespace ApiExamples {

/// <summary>
/// Shows how to rename merge fields in a Word document.
/// </summary>
class ExRenameMergeFields : public ApiExampleBase
{
    typedef ExRenameMergeFields ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Finds all merge fields in a Word document and changes their names.
    /// </summary>
    void Rename();
    
};

/// <summary>
/// Represents a facade object for a merge field in a Microsoft Word document.
/// </summary>
class MergeField : public System::Object
{
    typedef MergeField ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Gets or sets the name of the merge field.
    /// </summary>
    System::String get_Name();
    /// <summary>
    /// Gets or sets the name of the merge field.
    /// </summary>
    void set_Name(System::String value);
    
    MergeField(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart);
    
private:

    System::SharedPtr<Aspose::Words::Node> mFieldStart;
    System::SharedPtr<Aspose::Words::Node> mFieldSeparator;
    System::SharedPtr<Aspose::Words::Node> mFieldEnd;
    
    static System::SharedPtr<System::Text::RegularExpressions::Regex>& gRegex();
    void UpdateFieldCode(System::String fieldName);
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
} // namespace Words
} // namespace Aspose


