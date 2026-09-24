#pragma once

#include <xml/xpath/xpath_navigator.h>
#include <system/text/string_builder.h>
#include <system/string.h>
#include <gtest/gtest.h>
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
    //ExStart
    //ExFor:Node.NextSibling
    //ExFor:CompositeNode.FirstChild
    //ExFor:Node.IsComposite
    //ExFor:CompositeNode.IsComposite
    //ExFor:Node.NodeTypeToString
    //ExFor:Paragraph.NodeType
    //ExFor:Table.NodeType
    //ExFor:Node.NodeType
    //ExFor:Footnote.NodeType
    //ExFor:FormField.NodeType
    //ExFor:SmartTag.NodeType
    //ExFor:Cell.NodeType
    //ExFor:Row.NodeType
    //ExFor:Document.NodeType
    //ExFor:Comment.NodeType
    //ExFor:Run.NodeType
    //ExFor:Section.NodeType
    //ExFor:SpecialChar.NodeType
    //ExFor:Shape.NodeType
    //ExFor:FieldEnd.NodeType
    //ExFor:FieldSeparator.NodeType
    //ExFor:FieldStart.NodeType
    //ExFor:BookmarkStart.NodeType
    //ExFor:CommentRangeEnd.NodeType
    //ExFor:BuildingBlock.NodeType
    //ExFor:GlossaryDocument.NodeType
    //ExFor:BookmarkEnd.NodeType
    //ExFor:GroupShape.NodeType
    //ExFor:CommentRangeStart.NodeType
    //ExSummary:Shows how to traverse a composite node's tree of child nodes.
    void RecurseChildren();
    /// <summary>
    /// Recursively traverses a node tree while printing the type of each node
    /// with an indent depending on depth as well as the contents of all inline nodes.
    /// </summary>
    void TraverseAllNodes(System::SharedPtr<Aspose::Words::CompositeNode> parentNode, int32_t depth);
    //ExEnd
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
    //ExStart
    //ExFor:NodeChangingAction
    //ExFor:NodeChangingArgs.Action
    //ExFor:NodeChangingArgs.NewParent
    //ExFor:NodeChangingArgs.OldParent
    //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree in real-time as we edit it.
    void NodeChangingCallback();
    //ExEnd
    void NodeCollection();
    
protected:

    /// <summary>
    /// Traverses all children of a composite node and map the structure in the style of a directory tree.
    /// The amount of space indentation indicates depth relative to the initial node.
    /// Prints the text contents of the current node only if it is a Run.
    /// </summary>
    static void MapDocument(System::SharedPtr<System::Xml::XPath::XPathNavigator> navigator, System::SharedPtr<System::Text::StringBuilder> stringBuilder, int32_t depth);
    //ExEnd
    void TestNodeXPathNavigator(System::String navigatorResult, System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


