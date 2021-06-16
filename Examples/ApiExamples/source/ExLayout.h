#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Layout/IPageLayoutCallback.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/Layout/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Layout/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/PageLayoutCallbackArgs.h>
#include <Aspose.Words.Cpp/Layout/PageLayoutEvent.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/PageSet.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
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

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

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
        //ExSummary:Shows how to see the the ranges of pages that a node spans.
        auto doc = MakeObject<Document>();
        auto layoutCollector = MakeObject<LayoutCollector>(doc);

        // Call the "GetNumPagesSpanned" method to count how many pages the content of our document spans.
        // Since the document is empty, that number of pages is currently zero.
        ASPOSE_ASSERT_EQ(doc, layoutCollector->get_Document());
        ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));

        // Populate the document with 5 pages of content.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::SectionBreakEvenPage);
        builder->Write(u"Section 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertBreak(BreakType::PageBreak);

        // Before the layout collector, we need to call the "UpdatePageLayout" method to give us
        // an accurate figure for any layout-related metric, such as the page count.
        ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));

        layoutCollector->Clear();
        doc->UpdatePageLayout();

        ASSERT_EQ(5, layoutCollector->GetNumPagesSpanned(doc));

        // We can see the numbers of the start and end pages of any node and their overall page spans.
        SharedPtr<NodeCollection> nodes = doc->GetChildNodes(NodeType::Any, true);
        for (const auto& node : System::IterateOver(nodes))
        {
            std::cout << String::Format(u"->  NodeType.{0}: ", node->get_NodeType()) << std::endl;
            std::cout << (String::Format(u"\tStarts on page {0}, ends on page {1},", layoutCollector->GetStartPageIndex(node),
                                         layoutCollector->GetEndPageIndex(node)) +
                          String::Format(u" spanning {0} pages.", layoutCollector->GetNumPagesSpanned(node)))
                      << std::endl;
        }

        // We can iterate over the layout entities using a LayoutEnumerator.
        auto layoutEnumerator = MakeObject<LayoutEnumerator>(doc);

        ASSERT_EQ(LayoutEntityType::Page, layoutEnumerator->get_Type());

        // The LayoutEnumerator can traverse the collection of layout entities like a tree.
        // We can also apply it to any node's corresponding layout entity.
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
    //ExFor:Layout.LayoutEnumerator.MoveParent()
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
        // Open a document that contains a variety of layout entities.
        // Layout entities are pages, cells, rows, lines, and other objects included in the LayoutEntityType enum.
        // Each layout entity has a rectangular space that it occupies in the document body.
        auto doc = MakeObject<Document>(MyDir + u"Layout entities.docx");

        // Create an enumerator that can traverse these entities like a tree.
        auto layoutEnumerator = MakeObject<LayoutEnumerator>(doc);

        ASPOSE_ASSERT_EQ(doc, layoutEnumerator->get_Document());

        layoutEnumerator->MoveParent(LayoutEntityType::Page);

        ASSERT_EQ(LayoutEntityType::Page, layoutEnumerator->get_Type());
        ASSERT_THROW(layoutEnumerator->get_Text(), System::InvalidOperationException);

        // We can call this method to make sure that the enumerator will be at the first layout entity.
        layoutEnumerator->Reset();

        // There are two orders that determine how the layout enumerator continues traversing layout entities
        // when it encounters entities that span across multiple pages.
        // 1 -  In visual order:
        // When moving through an entity's children that span multiple pages,
        // page layout takes precedence, and we move to other child elements on this page and avoid the ones on the next.
        std::cout << "Traversing from first to last, elements between pages separated:" << std::endl;
        TraverseLayoutForward(layoutEnumerator, 1);

        // Our enumerator is now at the end of the collection. We can traverse the layout entities backwards to go back to the beginning.
        std::cout << "Traversing from last to first, elements between pages separated:" << std::endl;
        TraverseLayoutBackward(layoutEnumerator, 1);

        // 2 -  In logical order:
        // When moving through an entity's children that span multiple pages,
        // the enumerator will move between pages to traverse all the child entities.
        std::cout << "Traversing from first to last, elements between pages mixed:" << std::endl;
        TraverseLayoutForwardLogical(layoutEnumerator, 1);

        std::cout << "Traversing from last to first, elements between pages mixed:" << std::endl;
        TraverseLayoutBackwardLogical(layoutEnumerator, 1);
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
    /// in a depth-first manner, and in the "Visual" order.
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
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
    /// in a depth-first manner, and in the "Visual" order.
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
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
    /// in a depth-first manner, and in the "Logical" order.
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
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
    /// in a depth-first manner, and in the "Logical" order.
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
    /// Print information about layoutEnumerator's current entity to the console, while indenting the text with tab characters
    /// based on its depth relative to the root node that we provided in the constructor LayoutEnumerator instance.
    /// The rectangle that we process at the end represents the area and location that the entity takes up in the document.
    /// </summary>
    static void PrintCurrentEntity(SharedPtr<LayoutEnumerator> layoutEnumerator, int indent)
    {
        String tabs(u'\t', indent);

        std::cout << (layoutEnumerator->get_Kind() == String::Empty
                          ? String::Format(u"{0}-> Entity type: {1}", tabs, layoutEnumerator->get_Type())
                          : String::Format(u"{0}-> Entity type & kind: {1}, {2}", tabs, layoutEnumerator->get_Type(), layoutEnumerator->get_Kind()))
                  << std::endl;

        // Only spans can contain text.
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

    //ExStart
    //ExFor:IPageLayoutCallback
    //ExFor:IPageLayoutCallback.Notify(PageLayoutCallbackArgs)
    //ExFor:PageLayoutCallbackArgs.Event
    //ExFor:PageLayoutCallbackArgs.Document
    //ExFor:PageLayoutCallbackArgs.PageIndex
    //ExFor:PageLayoutEvent
    //ExSummary:Shows how to track layout changes with a layout callback.
    void PageLayoutCallback()
    {
        auto doc = MakeObject<Document>();
        doc->get_BuiltInDocumentProperties()->set_Title(u"My Document");

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        doc->get_LayoutOptions()->set_Callback(MakeObject<ExLayout::RenderPageLayoutCallback>());
        doc->UpdatePageLayout();

        doc->Save(ArtifactsDir + u"Layout.PageLayoutCallback.pdf");
    }

    /// <summary>
    /// Notifies us when we save the document to a fixed page format
    /// and renders a page that we perform a page reflow on to an image in the local file system.
    /// </summary>
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

            case PageLayoutEvent::ConversionFinished:
                NotifyConversionFinished(a);
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
            std::cout << "Part at page " << (a->get_PageIndex() + 1) << " reflow." << std::endl;
            RenderPage(a, a->get_PageIndex());
        }

        void NotifyConversionFinished(SharedPtr<PageLayoutCallbackArgs> a)
        {
            std::cout << "Document \"" << a->get_Document()->get_BuiltInDocumentProperties()->get_Title() << "\" converted to page format." << std::endl;
        }

        void RenderPage(SharedPtr<PageLayoutCallbackArgs> a, int pageIndex)
        {
            auto saveOptions = MakeObject<ImageSaveOptions>(SaveFormat::Png);
            saveOptions->set_PageSet(MakeObject<PageSet>(pageIndex));

            {
                auto stream = MakeObject<System::IO::FileStream>(ArtifactsDir + String::Format(u"PageLayoutCallback.page-{0} {1}.png", pageIndex + 1, ++mNum),
                                                                 System::IO::FileMode::Create);
                a->get_Document()->Save(stream, saveOptions);
            }
        }
    };
    //ExEnd
};

} // namespace ApiExamples
