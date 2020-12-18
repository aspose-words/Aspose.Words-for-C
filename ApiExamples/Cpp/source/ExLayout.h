#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Layout/Model/Callback/IPageLayoutCallback.h>
#include <Aspose.Words.Cpp/Layout/Model/Callback/PageLayoutCallbackArgs.h>
#include <Aspose.Words.Cpp/Layout/Model/Callback/PageLayoutEvent.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <drawing/rectangle_f.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExLayout : public ApiExampleBase
{
public:
    void LayoutCollector_()
    {
        //ExStart
        //ExFor:Layout.LayoutCollector
        //ExFor:Layout.LayoutCollector.#ctor(Document)
        //ExFor:Layout.LayoutCollector.Clear
        //ExFor:Layout.LayoutCollector.Document
        //ExFor:Layout.LayoutCollector.GetEndPageIndex(Node)
        //ExFor:Layout.LayoutCollector.GetEntity(Node)
        //ExFor:Layout.LayoutCollector.GetNumPagesSpanned(Node)
        //ExFor:Layout.LayoutCollector.GetStartPageIndex(Node)
        //ExFor:Layout.LayoutEnumerator.Current
        //ExSummary:Shows how to see the page spans of nodes.
        // Open a blank document and create a DocumentBuilder
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a LayoutCollector object for our document that will have information about the nodes we placed
        auto layoutCollector = MakeObject<LayoutCollector>(doc);

        // The document itself is a node that contains everything, which currently spans 0 pages
        ASPOSE_ASSERT_EQ(doc, layoutCollector->get_Document());
        ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));

        // Populate the document with sections and page breaks
        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        doc->AppendChild(MakeObject<Section>(doc));
        doc->get_LastSection()->AppendChild(MakeObject<Body>(doc));
        builder->MoveToDocumentEnd();
        builder->Write(u"Section 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);

        // The collected layout data won't automatically keep up with the real document contents
        ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));

        // After we clear the layout collection and update it, the layout entity collection will be populated with up-to-date information about our nodes
        // The page span for the document now shows 5, which is what we would expect after placing 4 page breaks
        layoutCollector->Clear();
        doc->UpdatePageLayout();
        ASSERT_EQ(5, layoutCollector->GetNumPagesSpanned(doc));

        // We can also see the start/end pages of any other node, and their overall page spans
        SharedPtr<NodeCollection> nodes = doc->GetChildNodes(NodeType::Any, true);
        for (auto node : System::IterateOver(nodes))
        {
            std::cout << String::Format(u"->  NodeType.{0}: ", node->get_NodeType()) << std::endl;
            std::cout << (String::Format(u"\tStarts on page {0}, ends on page {1},", layoutCollector->GetStartPageIndex(node),
                                                 layoutCollector->GetEndPageIndex(node)) +
                          String::Format(u" spanning {0} pages.", layoutCollector->GetNumPagesSpanned(node)))
                      << std::endl;
        }

        // We can iterate over the layout entities using a LayoutEnumerator
        auto layoutEnumerator = MakeObject<LayoutEnumerator>(doc);
        ASSERT_EQ(LayoutEntityType::Page, layoutEnumerator->get_Type());

        // The LayoutEnumerator can traverse the collection of layout entities like a tree
        // We can also point it to any node's corresponding layout entity like this
        layoutEnumerator->set_Current(layoutCollector->GetEntity(doc->GetChild(NodeType::Paragraph, 1, true)));
        ASSERT_EQ(LayoutEntityType::Span, layoutEnumerator->get_Type());
        ASSERT_EQ(u"¶", layoutEnumerator->get_Text());
        //ExEnd
    }

    //ExStart
    //ExFor:Layout.LayoutEntityType
    //ExFor:Layout.LayoutEnumerator
    //ExFor:Layout.LayoutEnumerator.#ctor(Document)
    //ExFor:Layout.LayoutEnumerator.Document
    //ExFor:Layout.LayoutEnumerator.Kind
    //ExFor:Layout.LayoutEnumerator.MoveFirstChild
    //ExFor:Layout.LayoutEnumerator.MoveLastChild
    //ExFor:Layout.LayoutEnumerator.MoveNext
    //ExFor:Layout.LayoutEnumerator.MoveNextLogical
    //ExFor:Layout.LayoutEnumerator.MoveParent
    //ExFor:Layout.LayoutEnumerator.MoveParent(Layout.LayoutEntityType)
    //ExFor:Layout.LayoutEnumerator.MovePrevious
    //ExFor:Layout.LayoutEnumerator.MovePreviousLogical
    //ExFor:Layout.LayoutEnumerator.PageIndex
    //ExFor:Layout.LayoutEnumerator.Rectangle
    //ExFor:Layout.LayoutEnumerator.Reset
    //ExFor:Layout.LayoutEnumerator.Text
    //ExFor:Layout.LayoutEnumerator.Type
    //ExSummary:Shows ways of traversing a document's layout entities.
    void LayoutEnumerator_()
    {
        // Open a document that contains a variety of layout entities
        // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
        // They are defined visually by the rectangular space that they occupy in the document
        auto doc = MakeObject<Document>(MyDir + u"Layout entities.docx");

        // Create an enumerator that can traverse these entities like a tree
        auto layoutEnumerator = MakeObject<LayoutEnumerator>(doc);
        ASPOSE_ASSERT_EQ(doc, layoutEnumerator->get_Document());

        layoutEnumerator->MoveParent(LayoutEntityType::Page);
        ASSERT_EQ(LayoutEntityType::Page, layoutEnumerator->get_Type());

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<void()> _local_func_0 = [&layoutEnumerator]()
        {
            std::cout << layoutEnumerator->get_Text() << std::endl;
        };

        ASSERT_THROW(_local_func_0(), System::InvalidOperationException);

        // We can call this method to make sure that the enumerator points to the very first entity before we go through it forwards
        layoutEnumerator->Reset();

        // "Visual order" means when moving through an entity's children that are broken across pages,
        // page layout takes precedence and we avoid elements in other pages and move to others on the same page
        std::cout << "Traversing from first to last, elements between pages separated:" << std::endl;
        TraverseLayoutForward(layoutEnumerator, 1);

        // Our enumerator is conveniently at the end of the collection for us to go through the collection backwards
        std::cout << "Traversing from last to first, elements between pages separated:" << std::endl;
        TraverseLayoutBackward(layoutEnumerator, 1);

        // "Logical order" means when moving through an entity's children that are broken across pages,
        // node relationships take precedence
        std::cout << "Traversing from first to last, elements between pages mixed:" << std::endl;
        TraverseLayoutForwardLogical(layoutEnumerator, 1);

        std::cout << "Traversing from last to first, elements between pages mixed:" << std::endl;
        TraverseLayoutBackwardLogical(layoutEnumerator, 1);
    }

protected:
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order.
    /// </summary>
    static void TraverseLayoutForward(SharedPtr<LayoutEnumerator> layoutEnumerator, int depth)
    {
        do
        {
            PrintCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator->MoveFirstChild())
            {
                TraverseLayoutForward(layoutEnumerator, depth + 1);
                layoutEnumerator->MoveParent();
            }
        } while (layoutEnumerator->MoveNext());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order.
    /// </summary>
    static void TraverseLayoutBackward(SharedPtr<LayoutEnumerator> layoutEnumerator, int depth)
    {
        do
        {
            PrintCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator->MoveLastChild())
            {
                TraverseLayoutBackward(layoutEnumerator, depth + 1);
                layoutEnumerator->MoveParent();
            }
        } while (layoutEnumerator->MovePrevious());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order.
    /// </summary>
    static void TraverseLayoutForwardLogical(SharedPtr<LayoutEnumerator> layoutEnumerator, int depth)
    {
        do
        {
            PrintCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator->MoveFirstChild())
            {
                TraverseLayoutForwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator->MoveParent();
            }
        } while (layoutEnumerator->MoveNextLogical());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order.
    /// </summary>
    static void TraverseLayoutBackwardLogical(SharedPtr<LayoutEnumerator> layoutEnumerator, int depth)
    {
        do
        {
            PrintCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator->MoveLastChild())
            {
                TraverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator->MoveParent();
            }
        } while (layoutEnumerator->MovePreviousLogical());
    }

    /// <summary>
    /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent.
    /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document.
    /// </summary>
    static void PrintCurrentEntity(SharedPtr<LayoutEnumerator> layoutEnumerator, int indent)
    {
        String tabs(u'\t', indent);

        std::cout << (layoutEnumerator->get_Kind() == String::Empty
                          ? String::Format(u"{0}-> Entity type: {1}", tabs, layoutEnumerator->get_Type())
                          : String::Format(u"{0}-> Entity type & kind: {1}, {2}", tabs, layoutEnumerator->get_Type(), layoutEnumerator->get_Kind()))
                  << std::endl;

        // Only spans can contain text
        if (layoutEnumerator->get_Type() == LayoutEntityType::Span)
        {
            std::cout << tabs << "   Span contents: \"" << layoutEnumerator->get_Text() << "\"" << std::endl;
        }

        System::Drawing::RectangleF leRect = layoutEnumerator->get_Rectangle();
        std::cout << tabs << "   Rectangle dimensions " << leRect.get_Width() << "x" << leRect.get_Height() << ", X=" << leRect.get_X()
                  << " Y=" << leRect.get_Y() << std::endl;
        std::cout << tabs << "   Page " << layoutEnumerator->get_PageIndex() << std::endl;
    }
    //ExEnd

public:
    //ExStart
    //ExFor:IPageLayoutCallback
    //ExFor:IPageLayoutCallback.Notify(PageLayoutCallbackArgs)
    //ExFor:PageLayoutCallbackArgs.Event
    //ExFor:PageLayoutCallbackArgs.Document
    //ExFor:PageLayoutCallbackArgs.PageIndex
    //ExFor:PageLayoutEvent
    //ExSummary:Shows how to track layout/rendering progress with layout callback.
    void PageLayoutCallback()
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        doc->get_LayoutOptions()->set_Callback(MakeObject<ExLayout::RenderPageLayoutCallback>());
        doc->UpdatePageLayout();
    }

private:
    class RenderPageLayoutCallback : public IPageLayoutCallback
    {
    public:
        void Notify(SharedPtr<PageLayoutCallbackArgs> a) override
        {
            switch (a->get_Event())
            {
            case PageLayoutEvent::PartReflowFinished:
                NotifyPartFinished(a);
                break;

            default:
                break;
            }
        }

        RenderPageLayoutCallback() : mNum(0)
        {
        }

    private:
        int mNum;

        void NotifyPartFinished(SharedPtr<PageLayoutCallbackArgs> a)
        {
            std::cout << "Part at page " << (a->get_PageIndex() + 1) << " reflow" << std::endl;
            RenderPage(a, a->get_PageIndex());
        }

        void RenderPage(SharedPtr<PageLayoutCallbackArgs> a, int pageIndex)
        {
            auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
            saveOptions->set_PageIndex(pageIndex);
            saveOptions->set_PageCount(1);

            {
                auto stream = MakeObject<System::IO::FileStream>(
                    ArtifactsDir + String::Format(u"PageLayoutCallback.page-{0} {1}.png", pageIndex + 1, ++mNum), System::IO::FileMode::Create);
                a->get_Document()->Save(stream, saveOptions);
            }
        }
    };
    //ExEnd
};

} // namespace ApiExamples
