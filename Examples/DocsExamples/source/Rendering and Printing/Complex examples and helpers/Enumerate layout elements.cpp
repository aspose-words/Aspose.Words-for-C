#include "Enumerate layout elements.h"

#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Rendering/PageInfo.h>
#include <drawing/bitmap.h>
#include <drawing/color.h>
#include <drawing/pens.h>
#include <drawing/rectangle.h>
#include <drawing/rectangle_f.h>
#include <drawing/size.h>
#include <drawing/size_f.h>
#include <system/convert.h>
#include <system/details/dispose_guard.h>
#include <system/object_ext.h>

#include "Rendering and Printing/Complex examples and helpers/Enumerate layout elements.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Rendering;
namespace DocsExamples { namespace Complex_examples_and_helpers {

namespace gtest_test {

class EnumerateLayoutElements : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Complex_examples_and_helpers::EnumerateLayoutElements> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Complex_examples_and_helpers::EnumerateLayoutElements>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Complex_examples_and_helpers::EnumerateLayoutElements> EnumerateLayoutElements::s_instance;

} // namespace gtest_test

void EnumerateLayoutElements::GetLayoutElements()
{
    auto doc = System::MakeObject<Document>(MyDir + u"Document layout.docx");

    // Enumerator which is used to "walk" the elements of a rendered document.
    auto layoutEnumerator = System::MakeObject<LayoutEnumerator>(doc);

    // Use the enumerator to write information about each layout element to the console.
    LayoutInfoWriter::Run(layoutEnumerator);

    // Adds a border around each layout element and saves each page as a JPEG image to the data directory.
    OutlineLayoutEntitiesRenderer::Run(doc, layoutEnumerator, ArtifactsDir);
}

namespace gtest_test {

TEST_F(EnumerateLayoutElements, GetLayoutElements)
{
    s_instance->GetLayoutElements();
}

} // namespace gtest_test

void LayoutInfoWriter::Run(System::SharedPtr<LayoutEnumerator> layoutEnumerator)
{
    DisplayLayoutElements(layoutEnumerator, System::String::Empty);
}

void LayoutInfoWriter::DisplayLayoutElements(System::SharedPtr<LayoutEnumerator> layoutEnumerator, System::String padding)
{
    do
    {
        DisplayEntityInfo(layoutEnumerator, padding);

        if (layoutEnumerator->MoveFirstChild())
        {
            // Recurse into this child element.
            DisplayLayoutElements(layoutEnumerator, AddPadding(padding));
            layoutEnumerator->MoveParent();
        }
    } while (layoutEnumerator->MoveNext());
}

void LayoutInfoWriter::DisplayEntityInfo(System::SharedPtr<LayoutEnumerator> layoutEnumerator, System::String padding)
{
    std::cout << (padding + System::ObjectExt::ToString(layoutEnumerator->get_Type()) + u" - " + layoutEnumerator->get_Kind());

    if (layoutEnumerator->get_Type() == LayoutEntityType::Span)
    {
        std::cout << (System::String(u" - ") + layoutEnumerator->get_Text());
    }

    std::cout << std::endl;
}

System::String LayoutInfoWriter::AddPadding(System::String padding)
{
    return padding + System::String(u' ', 4);
}

void OutlineLayoutEntitiesRenderer::Run(System::SharedPtr<Document> doc, System::SharedPtr<LayoutEnumerator> layoutEnumerator, System::String folderPath)
{
    // Make sure the enumerator is at the beginning of the document.
    layoutEnumerator->Reset();

    for (int32_t pageIndex = 0; pageIndex < doc->get_PageCount(); pageIndex++)
    {
        // Use the document class to find information about the current page.
        System::SharedPtr<PageInfo> pageInfo = doc->GetPageInfo(pageIndex);

        const float resolution = 150.0f;
        System::Drawing::Size pageSize = pageInfo->GetSizeInPixels(1.0f, resolution);

        {
            auto img = System::MakeObject<System::Drawing::Bitmap>(pageSize.get_Width(), pageSize.get_Height());
            img->SetResolution(resolution, resolution);

            {
                System::SharedPtr<System::Drawing::Graphics> g = System::Drawing::Graphics::FromImage(img);
                // Make the background white.
                g->Clear(System::Drawing::Color::get_White());

                // Render the page to the graphics.
                doc->RenderToScale(pageIndex, g, 0.0f, 0.0f, 1.0f);

                // Add an outline around each element on the page using the graphics object.
                AddBoundingBoxToElementsOnPage(layoutEnumerator, g);

                // Move the enumerator to the next page if there is one.
                layoutEnumerator->MoveNext();

                img->Save(folderPath + System::String::Format(u"EnumerateLayoutElements.Page_{0}.png", pageIndex + 1));
            }
        }
    }
}

void OutlineLayoutEntitiesRenderer::AddBoundingBoxToElementsOnPage(System::SharedPtr<LayoutEnumerator> layoutEnumerator,
                                                                   System::SharedPtr<System::Drawing::Graphics> g)
{
    do
    {
        // Use MoveLastChild and MovePrevious to enumerate from last to the first enumeration is done backward,
        // so the lines of child entities are drawn first and don't overlap the parent's lines.
        if (layoutEnumerator->MoveLastChild())
        {
            AddBoundingBoxToElementsOnPage(layoutEnumerator, g);
            layoutEnumerator->MoveParent();
        }

        // Convert the rectangle representing the position of the layout entity on the page from points to pixels.
        System::Drawing::RectangleF rectF = layoutEnumerator->get_Rectangle();
        System::Drawing::Rectangle rect(PointToPixel(rectF.get_Left(), g->get_DpiX()), PointToPixel(rectF.get_Top(), g->get_DpiY()),
                                        PointToPixel(rectF.get_Width(), g->get_DpiX()), PointToPixel(rectF.get_Height(), g->get_DpiY()));

        // Draw a line around the layout entity on the page.
        g->DrawRectangle(GetColoredPenFromType(layoutEnumerator->get_Type()), rect);

        // Stop after all elements on the page have been processed.
        if (layoutEnumerator->get_Type() == LayoutEntityType::Page)
        {
            return;
        }
    } while (layoutEnumerator->MovePrevious());
}

System::SharedPtr<System::Drawing::Pen> OutlineLayoutEntitiesRenderer::GetColoredPenFromType(LayoutEntityType type)
{
    switch (type)
    {
    case LayoutEntityType::Cell:
        return System::Drawing::Pens::get_Purple();

    case LayoutEntityType::Column:
        return System::Drawing::Pens::get_Green();

    case LayoutEntityType::Comment:
        return System::Drawing::Pens::get_LightBlue();

    case LayoutEntityType::Endnote:
        return System::Drawing::Pens::get_DarkRed();

    case LayoutEntityType::Footnote:
        return System::Drawing::Pens::get_DarkBlue();

    case LayoutEntityType::HeaderFooter:
        return System::Drawing::Pens::get_DarkGreen();

    case LayoutEntityType::Line:
        return System::Drawing::Pens::get_Blue();

    case LayoutEntityType::NoteSeparator:
        return System::Drawing::Pens::get_LightGreen();

    case LayoutEntityType::Page:
        return System::Drawing::Pens::get_Red();

    case LayoutEntityType::Row:
        return System::Drawing::Pens::get_Orange();

    case LayoutEntityType::Span:
        return System::Drawing::Pens::get_Red();

    case LayoutEntityType::TextBox:
        return System::Drawing::Pens::get_Yellow();

    default:
        return System::Drawing::Pens::get_Red();
    }
}

int32_t OutlineLayoutEntitiesRenderer::PointToPixel(float value, double resolution)
{
    return System::Convert::ToInt32(ConvertUtil::PointToPixel(value, resolution));
}

}} // namespace DocsExamples::Complex_examples_and_helpers
