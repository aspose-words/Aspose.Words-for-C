#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Nodes/INodeChangingCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class CompositeNode; } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { class NodeChangingArgs; } }

namespace ApiExamples {

class ExNode : public ApiExampleBase
{
private:

    /// <summary>
    /// Prints every node insertion/removal as it takes place in the document.
    /// </summary>
    class NodeChangingPrinter : public Aspose::Words::INodeChangingCallback
    {
        typedef NodeChangingPrinter ThisType;
        typedef Aspose::Words::INodeChangingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void NodeInserting(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeInserted(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoving(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoved(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        
    };
    
    
public:

    void CloneCompositeNode();
    void GetParentNode();
    void OwnerDocument();
    void EnumerateChildNodes();
    void IndexChildNodes();
    void RecurseAllNodes();
    /// <summary>
    /// Recursively traverses a node tree while printing the type of each node with an indent depending on depth as well as the contents of all inline nodes.
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

    void TestNodeXPathNavigator(System::String navigatorResult, System::SharedPtr<Aspose::Words::Document> doc);
    
};

} // namespace ApiExamples


