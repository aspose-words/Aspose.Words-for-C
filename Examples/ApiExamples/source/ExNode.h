#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <xml/xpath/xpath_navigator.h>
#include <system/text/string_builder.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Nodes/NodeChangingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/INodeChangingCallback.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExNode : public ApiExampleBase
{
    typedef ExNode ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Prints every node insertion/removal as it takes place in the document.
    /// </summary>
    class NodeChangingPrinter : public INodeChangingCallback
    {
        typedef NodeChangingPrinter ThisType;
        typedef INodeChangingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    private:
    
        void NodeInserting(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeInserted(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoving(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoved(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        
    };
    
    
public:

    void CloneCompositeNode();
    void GetParentNode();
    void OwnerDocument();
    void ChildNodesEnumerate();
    void RecurseChildren();
    /// <summary>
    /// Recursively traverses a node tree while printing the type of each node
    /// with an indent depending on depth as well as the contents of all inline nodes.
    /// </summary>
    void TraverseAllNodes(System::SharedPtr<Aspose::Words::CompositeNode> parentNode, int32_t depth);
    void RemoveNodes();
    void EnumNextSibling();
    void TypedAccess();
    void RemoveChild();
    void CreateAndAddParagraphNode();
    void RemoveSmartTagsFromCompositeNode();
    void GetIndexOfNode();
    void ConvertNodeToHtmlWithDefaultOptions();
    void TypedNodeCollectionToArray();
    void NodeEnumerationHotRemove();
    void NodeChangingCallback();
    void NodeCollection();
    
protected:

    /// <summary>
    /// Traverses all children of a composite node and map the structure in the style of a directory tree.
    /// The amount of space indentation indicates depth relative to the initial node.
    /// Prints the text contents of the current node only if it is a Run.
    /// </summary>
    static void MapDocument(System::SharedPtr<System::Xml::XPath::XPathNavigator> navigator, System::SharedPtr<System::Text::StringBuilder> stringBuilder, int32_t depth);
    void TestNodeXPathNavigator(System::String navigatorResult, System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


