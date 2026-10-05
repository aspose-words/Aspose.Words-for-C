#pragma once

#include <system/shared_ptr.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Layout/Model/Callback/PageLayoutCallbackArgs.h>
#include <Aspose.Words.Cpp/Layout/Model/Callback/IPageLayoutCallback.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Layout;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExLayout : public ApiExampleBase
{
    typedef ExLayout ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Notifies us when we save the document to a fixed page format
    /// and renders a page that we perform a page reflow on to an image in the local file system.
    /// </summary>
    class RenderPageLayoutCallback : public IPageLayoutCallback
    {
        typedef RenderPageLayoutCallback ThisType;
        typedef IPageLayoutCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void Notify(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a) override;
        
        RenderPageLayoutCallback();
        
    private:
    
        int32_t mNum;
        
        void NotifyPartFinished(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a);
        void NotifyConversionFinished(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a);
        void RenderPage(System::SharedPtr<Aspose::Words::Layout::PageLayoutCallbackArgs> a, int32_t pageIndex);
        
    };
    
    
public:

    void LayoutCollector();
    //ExStart
    //ExFor:LayoutEntityType
    //ExFor:LayoutEnumerator
    //ExFor:LayoutEnumerator.#ctor(Document)
    //ExFor:LayoutEnumerator.Document
    //ExFor:LayoutEnumerator.Kind
    //ExFor:LayoutEnumerator.MoveFirstChild
    //ExFor:LayoutEnumerator.MoveLastChild
    //ExFor:LayoutEnumerator.MoveNext
    //ExFor:LayoutEnumerator.MoveNextLogical
    //ExFor:LayoutEnumerator.MoveParent
    //ExFor:LayoutEnumerator.MoveParent(LayoutEntityType)
    //ExFor:LayoutEnumerator.MovePrevious
    //ExFor:LayoutEnumerator.MovePreviousLogical
    //ExFor:LayoutEnumerator.PageIndex
    //ExFor:LayoutEnumerator.Rectangle
    //ExFor:LayoutEnumerator.Reset
    //ExFor:LayoutEnumerator.Text
    //ExFor:LayoutEnumerator.Type
    //ExSummary:Shows ways of traversing a document's layout entities.
    void LayoutEnumerator();
    //ExEnd
    //ExStart
    //ExFor:IPageLayoutCallback
    //ExFor:IPageLayoutCallback.Notify(PageLayoutCallbackArgs)
    //ExFor:PageLayoutCallbackArgs
    //ExFor:PageLayoutCallbackArgs.Event
    //ExFor:PageLayoutCallbackArgs.Document
    //ExFor:PageLayoutCallbackArgs.PageIndex
    //ExFor:PageLayoutEvent
    //ExFor:LayoutOptions.Callback
    //ExSummary:Shows how to track layout changes with a layout callback.
    void PageLayoutCallback();
    //ExEnd
    void RestartPageNumberingInContinuousSection();
    
protected:

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
    /// in a depth-first manner, and in the "Visual" order.
    /// </summary>
    static void TraverseLayoutForward(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
    /// in a depth-first manner, and in the "Visual" order.
    /// </summary>
    static void TraverseLayoutBackward(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
    /// in a depth-first manner, and in the "Logical" order.
    /// </summary>
    static void TraverseLayoutForwardLogical(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
    /// in a depth-first manner, and in the "Logical" order.
    /// </summary>
    static void TraverseLayoutBackwardLogical(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Print information about layoutEnumerator's current entity to the console, while indenting the text with tab characters
    /// based on its depth relative to the root node that we provided in the constructor LayoutEnumerator instance.
    /// The rectangle that we process at the end represents the area and location that the entity takes up in the document.
    /// </summary>
    static void PrintCurrentEntity(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t indent);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


