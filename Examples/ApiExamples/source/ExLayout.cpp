// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExLayout.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/rectangle_f.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageSet.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>
#include <Aspose.Words.Cpp/Layout/Public/ContinuousSectionRestart.h>
#include <Aspose.Words.Cpp/Layout/Model/Callback/PageLayoutEvent.h>


using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1893098702u, ::Aspose::Words::ApiExamples::ExLayout::RenderPageLayoutCallback, ThisTypeBaseTypesInfo);

void ExLayout::RenderPageLayoutCallback::Notify(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a)
{
    switch (a->get_Event())
    {
        case Aspose::Words::Layout::PageLayoutEvent::PartReflowFinished:
            NotifyPartFinished(a);
            break;
        
        case Aspose::Words::Layout::PageLayoutEvent::ConversionFinished:
            NotifyConversionFinished(a);
            break;
        
        default:
            break;
    }
}

void ExLayout::RenderPageLayoutCallback::NotifyPartFinished(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a)
{
    std::cout << System::String::Format(u"Part at page {0} reflow.", a->get_PageIndex() + 1) << std::endl;
    RenderPage(a, a->get_PageIndex());
}

void ExLayout::RenderPageLayoutCallback::NotifyConversionFinished(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a)
{
    std::cout << System::String::Format(u"Document \"{0}\" converted to page format.", a->get_Document()->get_BuiltInDocumentProperties()->get_Title()) << std::endl;
}

void ExLayout::RenderPageLayoutCallback::RenderPage(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a, int32_t pageIndex)
{
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    saveOptions->set_PageSet(System::MakeObject<Aspose::Words::Saving::PageSet>(pageIndex));
    
    {
        auto stream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + System::String::Format(u"PageLayoutCallback.page-{0} {1}.png", pageIndex + 1, ++mNum), System::IO::FileMode::Create);
        a->get_Document()->Save(stream, saveOptions);
    }
}

ExLayout::RenderPageLayoutCallback::RenderPageLayoutCallback() : mNum(0)
{
}


RTTI_INFO_IMPL_HASH(3498384122u, ::Aspose::Words::ApiExamples::ExLayout, ThisTypeBaseTypesInfo);

void ExLayout::TraverseLayoutForward(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth)
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

void ExLayout::TraverseLayoutBackward(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth)
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

void ExLayout::TraverseLayoutForwardLogical(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth)
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

void ExLayout::TraverseLayoutBackwardLogical(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth)
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

void ExLayout::PrintCurrentEntity(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t indent)
{
    System::String tabs(u'\t', indent);
    
    std::cout << (layoutEnumerator->get_Kind() == System::String::Empty ? System::String::Format(u"{0}-> Entity type: {1}", tabs, layoutEnumerator->get_Type()) : System::String::Format(u"{0}-> Entity type & kind: {1}, {2}", tabs, layoutEnumerator->get_Type(), layoutEnumerator->get_Kind())) << std::endl;
    
    // Only spans can contain text.
    if (layoutEnumerator->get_Type() == Aspose::Words::Layout::LayoutEntityType::Span)
    {
        std::cout << System::String::Format(u"{0}   Span contents: \"{1}\"", tabs, layoutEnumerator->get_Text()) << std::endl;
    }
    
    System::Drawing::RectangleF leRect = layoutEnumerator->get_Rectangle();
    std::cout << System::String::Format(u"{0}   Rectangle dimensions {1}x{2}, X={3} Y={4}", tabs, leRect.get_Width(), leRect.get_Height(), leRect.get_X(), leRect.get_Y()) << std::endl;
    std::cout << System::String::Format(u"{0}   Page {1}", tabs, layoutEnumerator->get_PageIndex()) << std::endl;
}


namespace gtest_test
{

class ExLayout : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExLayout> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExLayout>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExLayout> ExLayout::s_instance;

} // namespace gtest_test

void ExLayout::LayoutCollector()
{
    //ExStart
    //ExFor:LayoutCollector
    //ExFor:LayoutCollector.#ctor(Document)
    //ExFor:LayoutCollector.Clear
    //ExFor:LayoutCollector.Document
    //ExFor:LayoutCollector.GetEndPageIndex(Node)
    //ExFor:LayoutCollector.GetEntity(Node)
    //ExFor:LayoutCollector.GetNumPagesSpanned(Node)
    //ExFor:LayoutCollector.GetStartPageIndex(Node)
    //ExFor:LayoutEnumerator.Current
    //ExSummary:Shows how to see the the ranges of pages that a node spans.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto layoutCollector = System::MakeObject<Aspose::Words::Layout::LayoutCollector>(doc);
    
    // Call the "GetNumPagesSpanned" method to count how many pages the content of our document spans.
    // Since the document is empty, that number of pages is currently zero.
    ASPOSE_ASSERT_EQ(doc, layoutCollector->get_Document());
    ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));
    
    // Populate the document with 5 pages of content.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakEvenPage);
    builder->Write(u"Section 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    
    // Before the layout collector, we need to call the "UpdatePageLayout" method to give us
    // an accurate figure for any layout-related metric, such as the page count.
    ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));
    
    layoutCollector->Clear();
    doc->UpdatePageLayout();
    
    ASSERT_EQ(5, layoutCollector->GetNumPagesSpanned(doc));
    
    // We can see the numbers of the start and end pages of any node and their overall page spans.
    System::SharedPtr<Aspose::Words::NodeCollection> nodes = doc->GetChildNodes(Aspose::Words::NodeType::Any, true);
    for (auto&& node : System::IterateOver(nodes))
    {
        std::cout << System::String::Format(u"->  NodeType.{0}: ", node->get_NodeType()) << std::endl;
        std::cout << (System::String::Format(u"\tStarts on page {0}, ends on page {1},", layoutCollector->GetStartPageIndex(node), layoutCollector->GetEndPageIndex(node)) + System::String::Format(u" spanning {0} pages.", layoutCollector->GetNumPagesSpanned(node))) << std::endl;
    }
    
    // We can iterate over the layout entities using a LayoutEnumerator.
    auto layoutEnumerator = System::MakeObject<Aspose::Words::Layout::LayoutEnumerator>(doc);
    
    ASSERT_EQ(Aspose::Words::Layout::LayoutEntityType::Page, layoutEnumerator->get_Type());
    
    // The LayoutEnumerator can traverse the collection of layout entities like a tree.
    // We can also apply it to any node's corresponding layout entity.
    layoutEnumerator->set_Current(layoutCollector->GetEntity(doc->GetChild(Aspose::Words::NodeType::Paragraph, 1, true)));
    
    ASSERT_EQ(Aspose::Words::Layout::LayoutEntityType::Span, layoutEnumerator->get_Type());
    ASSERT_EQ(u"¶", layoutEnumerator->get_Text());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLayout, LayoutCollector)
{
    s_instance->LayoutCollector();
}

} // namespace gtest_test

void ExLayout::LayoutEnumerator()
{
    // Open a document that contains a variety of layout entities.
    // Layout entities are pages, cells, rows, lines, and other objects included in the LayoutEntityType enum.
    // Each layout entity has a rectangular space that it occupies in the document body.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Layout entities.docx");
    
    // Create an enumerator that can traverse these entities like a tree.
    auto layoutEnumerator = System::MakeObject<Aspose::Words::Layout::LayoutEnumerator>(doc);
    
    ASPOSE_ASSERT_EQ(doc, layoutEnumerator->get_Document());
    
    layoutEnumerator->MoveParent(Aspose::Words::Layout::LayoutEntityType::Page);
    
    ASSERT_EQ(Aspose::Words::Layout::LayoutEntityType::Page, layoutEnumerator->get_Type());
    ASSERT_THROW(static_cast<std::function<void()>>([&layoutEnumerator]() -> void
    {
        std::cout << layoutEnumerator->get_Text() << std::endl;
    })(), System::InvalidOperationException);
    
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

namespace gtest_test
{

TEST_F(ExLayout, LayoutEnumerator)
{
    s_instance->LayoutEnumerator();
}

} // namespace gtest_test

void ExLayout::PageLayoutCallback()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->get_BuiltInDocumentProperties()->set_Title(u"My Document");
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    doc->get_LayoutOptions()->set_Callback(System::MakeObject<Aspose::Words::ApiExamples::ExLayout::RenderPageLayoutCallback>());
    doc->UpdatePageLayout();
    
    doc->Save(get_ArtifactsDir() + u"Layout.PageLayoutCallback.pdf");
}

namespace gtest_test
{

TEST_F(ExLayout, PageLayoutCallback)
{
    s_instance->PageLayoutCallback();
}

} // namespace gtest_test

void ExLayout::RestartPageNumberingInContinuousSection()
{
    //ExStart
    //ExFor:LayoutOptions.ContinuousSectionPageNumberingRestart
    //ExFor:ContinuousSectionRestart
    //ExSummary:Shows how to control page numbering in a continuous section.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Continuous section page numbering.docx");
    
    // By default Aspose.Words behavior matches the Microsoft Word 2019.
    // If you need old Aspose.Words behavior, repetitive Microsoft Word 2016, use 'ContinuousSectionRestart.FromNewPageOnly'.
    // Page numbering restarts only if there is no other content before the section on the page where the section starts,
    // because of that the numbering will reset to 2 from the second page.
    doc->get_LayoutOptions()->set_ContinuousSectionPageNumberingRestart(Aspose::Words::Layout::ContinuousSectionRestart::FromNewPageOnly);
    doc->UpdatePageLayout();
    
    doc->Save(get_ArtifactsDir() + u"Layout.RestartPageNumberingInContinuousSection.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLayout, RestartPageNumberingInContinuousSection)
{
    s_instance->RestartPageNumberingInContinuousSection();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
