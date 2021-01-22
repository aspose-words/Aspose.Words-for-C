#include "stdafx.h"
#include "examples.h"

#include <drawing/bitmap.h>
#include <drawing/pen.h>
#include <drawing/pens.h>
#include <drawing/graphics.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Rendering/PageInfo.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Rendering;

namespace
{
    class LayoutInfoWriter
    {
    public:
        static void Run(System::SharedPtr<LayoutEnumerator> it);

    private:
        LayoutInfoWriter();
        static void DisplayLayoutElements(System::SharedPtr<LayoutEnumerator> it, System::String const &padding);
        static void DisplayEntityInfo(System::SharedPtr<LayoutEnumerator> it, System::String const &padding);
        static System::String AddPadding(System::String const &padding);
    };

    void LayoutInfoWriter::Run(System::SharedPtr<LayoutEnumerator> it)
    {
        DisplayLayoutElements(it, System::String::Empty);
    }

    void LayoutInfoWriter::DisplayLayoutElements(System::SharedPtr<LayoutEnumerator> it, System::String const &padding)
    {
        do
        {
            DisplayEntityInfo(it, padding);

            if (it->MoveFirstChild())
            {
                // Recurse into this child element.
                DisplayLayoutElements(it, AddPadding(padding));
                it->MoveParent();
            }
        } while (it->MoveNext());
    }

    void LayoutInfoWriter::DisplayEntityInfo(System::SharedPtr<LayoutEnumerator> it, System::String const &padding)
    {
        std::cout << padding.ToUtf8String() << System::ObjectExt::ToString(it->get_Type()).ToUtf8String() << " - " << it->get_Kind().ToUtf8String();

        if (it->get_Type() == LayoutEntityType::Span)
        {
            std::cout << " - " << it->get_Text().ToUtf8String();
        }

        std::cout << std::endl;
    }

    System::String LayoutInfoWriter::AddPadding(System::String const &padding)
    {
        return padding + System::String(u' ', 4);
    }

    class OutlineLayoutEntitiesRenderer
    {
    public:
        static void Run(System::SharedPtr<Document> doc, System::SharedPtr<LayoutEnumerator> it, System::String const &folderPath);

    private:
        OutlineLayoutEntitiesRenderer();
        static void AddBoundingBoxToElementsOnPage(System::SharedPtr<LayoutEnumerator> it, System::SharedPtr<System::Drawing::Graphics> g);
        static System::SharedPtr<System::Drawing::Pen> GetColoredPenFromType(LayoutEntityType type);
        static int32_t PointToPixel(float value, double resolution);
    };

    void OutlineLayoutEntitiesRenderer::Run(System::SharedPtr<Document> doc, System::SharedPtr<LayoutEnumerator> it, System::String const &outputDataDir)
    {
        // Make sure the enumerator is at the beginning of the document.
        it->Reset();

        for (int32_t pageIndex = 0; pageIndex < doc->get_PageCount(); pageIndex++)
        {
            // Use the document class to find information about the current page.
            System::SharedPtr<PageInfo> pageInfo = doc->GetPageInfo(pageIndex);

            const float resolution = 150.0f;
            System::Drawing::Size pageSize = pageInfo->GetSizeInPixels(1.0f, resolution);

            {
                // Clearing resources under 'using' statement
                DisposableHolder<System::Drawing::Bitmap> imgHolder(System::MakeObject<System::Drawing::Bitmap>(pageSize.get_Width(), pageSize.get_Height()));

                imgHolder.GetObject()->SetResolution(resolution, resolution);

                {
                    // Clearing resources under 'using' statement
                    DisposableHolder<System::Drawing::Graphics> gHolder(System::Drawing::Graphics::FromImage(imgHolder.GetObject()));

                    // Make the background white.
                    gHolder.GetObject()->Clear(System::Drawing::Color::get_White());

                    // Render the page to the graphics.
                    doc->RenderToScale(pageIndex, gHolder.GetObject(), 0.0f, 0.0f, 1.0f);

                    // Add an outline around each element on the page using the graphics object.
                    AddBoundingBoxToElementsOnPage(it, gHolder.GetObject());

                    // Move the enumerator to the next page if there is one.
                    it->MoveNext();

                    imgHolder.GetObject()->Save(outputDataDir + System::String::Format(u"TestFile Page {0}.png", pageIndex + 1));
                }
            }
        }
    }

    void OutlineLayoutEntitiesRenderer::AddBoundingBoxToElementsOnPage(System::SharedPtr<LayoutEnumerator> it, System::SharedPtr<System::Drawing::Graphics> g)
    {
        do
        {
            // This time instead of MoveFirstChild and MoveNext, we use MoveLastChild and MovePrevious to enumerate from last to first.
            // Enumeration is done backward so the lines of child entities are drawn first and don't overlap the lines of the parent.
            if (it->MoveLastChild())
            {
                AddBoundingBoxToElementsOnPage(it, g);
                it->MoveParent();
            }

            // Convert the rectangle representing the position of the layout entity on the page from points to pixels.
            System::Drawing::RectangleF rectF = it->get_Rectangle();
            System::Drawing::Rectangle rect(PointToPixel(rectF.get_Left(), g->get_DpiX()),
                                            PointToPixel(rectF.get_Top(), g->get_DpiY()),
                                            PointToPixel(rectF.get_Width(), g->get_DpiX()),
                                            PointToPixel(rectF.get_Height(), g->get_DpiY()));

            // Draw a line around the layout entity on the page.
            g->DrawRectangle(GetColoredPenFromType(it->get_Type()), rect);

            // Stop after all elements on the page have been procesed.
            if (it->get_Type() == LayoutEntityType::Page)
            {
                return;
            }

        } while (it->MovePrevious());
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
}

void EnumerateLayoutElements()
{
    std::cout << "EnumerateLayoutElements example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.EnumerateLayout.docx");

    // This creates an enumerator which is used to "walk" the elements of a rendered document.
    System::SharedPtr<LayoutEnumerator> it = System::MakeObject<LayoutEnumerator>(doc);

    // This sample uses the enumerator to write information about each layout element to the console.
    LayoutInfoWriter::Run(it);

    // This sample adds a border around each layout element and saves each page as a JPEG image to the data directory.
    OutlineLayoutEntitiesRenderer::Run(doc, it, outputDataDir);

    std::cout << "Enumerate layout elements example ran successfully." << std::endl;
    std::cout << "EnumerateLayoutElements example finished." << std::endl << std::endl;
}