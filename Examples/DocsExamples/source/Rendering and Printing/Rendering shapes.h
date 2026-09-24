#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/InlineStory.h>
#include <Aspose.Words.Cpp/Layout/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Rendering/ShapeRenderer.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/ImageColorMode.h>
#include <Aspose.Words.Cpp/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Story.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <drawing/bitmap.h>
#include <drawing/color.h>
#include <drawing/graphics.h>
#include <drawing/imaging/image_format.h>
#include <drawing/point.h>
#include <drawing/rectangle.h>
#include <drawing/rectangle_f.h>
#include <drawing/size.h>
#include <system/details/dispose_guard.h>
#include <system/exceptions.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/math.h>
#include <system/enumerator_adapter.h>
#include <system/object_ext.h>
#include <system/primitive_types.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Rendering_and_Printing {

class RenderingShapes : public DocsExamplesBase
{
public:
    void RenderShapeAsEmf()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Retrieve the target shape from the document.
        auto shape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:RenderShapeAsEmf
        //GistId:010c69d7d7a0108c7852e317259376f9
        SharedPtr<ShapeRenderer> render = shape->GetShapeRenderer();

        auto imageOptions = MakeObject<ImageSaveOptions>(SaveFormat::Emf);
        imageOptions->set_Scale(1.5f);

        render->Save(ArtifactsDir + u"RenderShape.RenderShapeAsEmf.emf", imageOptions);
        //ExEnd:RenderShapeAsEmf
    }

    void RenderShapeAsJpeg()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto shape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:RenderShapeAsJpeg
        //GistId:010c69d7d7a0108c7852e317259376f9
        auto render = MakeObject<ShapeRenderer>(shape);

        auto imageOptions = MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);
        imageOptions->set_ImageColorMode(ImageColorMode::Grayscale);
        imageOptions->set_ImageBrightness(0.45f);

        {
            auto stream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"RenderShape.RenderShapeAsJpeg.jpg", System::IO::FileMode::Create);
            render->Save(stream, imageOptions);
        }
        //ExEnd:RenderShapeAsJpeg
    }

    //ExStart:RenderShapeToGraphics
    //GistId:010c69d7d7a0108c7852e317259376f9
    void RenderShapeToGraphics()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto shape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        SharedPtr<ShapeRenderer> render = shape->GetShapeRenderer();

        // Find the size that the shape will be rendered to at the specified scale and resolution.
        System::Drawing::Size shapeSizeInPixels = render->GetSizeInPixels(1.0f, 96.0f);

        // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
        // and make sure that the graphics canvas is large enough to compensate for this.
        int maxSide = System::Math::Max(shapeSizeInPixels.get_Width(), shapeSizeInPixels.get_Height());

        {
            auto image = MakeObject<System::Drawing::Bitmap>((int)(maxSide * 1.25), (int)(maxSide * 1.25));
            // Rendering to a graphics object means we can specify settings and transformations to be applied to the rendered shape.
            // In our case we will rotate the rendered shape.
            {
                SharedPtr<System::Drawing::Graphics> graphics = System::Drawing::Graphics::FromImage(image);
                // Clear the shape with the background color of the document.
                graphics->Clear(shape->get_Document()->get_PageColor());
                // Center the rotation using the translation method below.
                graphics->TranslateTransform((float)image->get_Width() / 8, (float)image->get_Height() / 2);
                // Rotate the image by 45 degrees.
                graphics->RotateTransform(45.0f);
                // Undo the translation.
                graphics->TranslateTransform(-(float)image->get_Width() / 8, -(float)image->get_Height() / 2);

                // Render the shape onto the graphics object.
                render->RenderToSize(graphics, 0.0f, 0.0f, static_cast<float>(shapeSizeInPixels.get_Width()),
                                     static_cast<float>(shapeSizeInPixels.get_Height()));
            }

            image->Save(ArtifactsDir + u"RenderShape.RenderShapeToGraphics.png", System::Drawing::Imaging::ImageFormat::get_Png());
        }
    }
    //ExEnd:RenderShapeToGraphics

    void RenderCellToImage()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        //ExStart:RenderCellToImage
        auto cell = System::ExplicitCast<Cell>(doc->GetChild(NodeType::Cell, 2, true));
        SharedPtr<Document> tmp = ConvertToImage(doc, cell);
        tmp->Save(ArtifactsDir + u"RenderShape.RenderCellToImage.png");
        //ExEnd:RenderCellToImage
    }

    void RenderRowToImage()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        //ExStart:RenderRowToImage
        auto row = System::ExplicitCast<Row>(doc->GetChild(NodeType::Row, 0, true));
        SharedPtr<Document> tmp = ConvertToImage(doc, row);
        tmp->Save(ArtifactsDir + u"RenderShape.RenderRowToImage.png");
        //ExEnd:RenderRowToImage
    }

    void RenderParagraphToImage()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        //ExStart:RenderParagraphToImage
        SharedPtr<Shape> textBoxShape = builder->InsertShape(ShapeType::TextBox, 150, 100);

        builder->MoveTo(textBoxShape->get_LastParagraph());
        builder->Write(u"Vertical text");

        auto options = MakeObject<ImageSaveOptions>(SaveFormat::Png);
        options->set_PaperColor(System::Drawing::Color::get_LightPink());

        SharedPtr<Document> tmp = ConvertToImage(doc, textBoxShape->get_LastParagraph());
        tmp->Save(ArtifactsDir + u"RenderShape.RenderParagraphToImage.png");
        //ExEnd:RenderParagraphToImage
    }

    void FindShapeSizes()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto shape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:FindShapeSizes
        //GistId:010c69d7d7a0108c7852e317259376f9
        System::Drawing::Size shapeRenderedSize = shape->GetShapeRenderer()->GetSizeInPixels(1.0f, 96.0f);

        {
            auto image = MakeObject<System::Drawing::Bitmap>(shapeRenderedSize.get_Width(), shapeRenderedSize.get_Height());
            {
                SharedPtr<System::Drawing::Graphics> graphics = System::Drawing::Graphics::FromImage(image);
                // Render shape onto the graphics object using the RenderToScale
                // or RenderToSize methods of ShapeRenderer class.
            }
        }
        //ExEnd:FindShapeSizes
    }

    void RenderShapeImage()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto shape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:RenderShapeImage
        //GistId:010c69d7d7a0108c7852e317259376f9
        shape->GetShapeRenderer()->Save(ArtifactsDir + u"RenderShape.RenderShapeImage.jpg", System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(SaveFormat::Jpeg));
        //ExEnd:RenderShapeImage
    }

    /// <summary>
    /// Renders any node in a document into an image.
    /// </summary>
    /// <param name="doc">The current document.</param>
    /// <param name="node">The node to render.</param>
    SharedPtr<Document> ConvertToImage(SharedPtr<Document> doc, SharedPtr<CompositeNode> node)
    {
        SharedPtr<Document> tmp = CreateTemporaryDocument(doc, node);
        AppendNodeContent(tmp, node);
        AdjustDocumentLayout(tmp);

        return tmp;
    }

    /// <summary>
    /// Creates a temporary document for further rendering.
    /// </summary>
    SharedPtr<Document> CreateTemporaryDocument(SharedPtr<Document> doc, SharedPtr<CompositeNode> node)
    {
        auto tmp = System::ExplicitCast<Document>(doc->Clone(false));
        tmp->get_Sections()->Add(tmp->ImportNode(node->GetAncestor(NodeType::Section), false, ImportFormatMode::UseDestinationStyles));
        tmp->get_FirstSection()->AppendChild(MakeObject<Body>(tmp));
        tmp->get_FirstSection()->get_PageSetup()->set_TopMargin(0);
        tmp->get_FirstSection()->get_PageSetup()->set_BottomMargin(0);

        return tmp;
    }

    /// <summary>
    /// Adds a node to a temporary document.
    /// </summary>
    void AppendNodeContent(SharedPtr<Document> tmp, SharedPtr<CompositeNode> node)
    {
        auto headerFooter = System::AsCast<HeaderFooter>(node);
        if (headerFooter != nullptr)
        {
            for (const auto& hfNode : System::IterateOver(headerFooter->GetChildNodes(NodeType::Any, false)))
            {
                tmp->get_FirstSection()->get_Body()->AppendChild(tmp->ImportNode(hfNode, true, ImportFormatMode::UseDestinationStyles));
            }
        }
        else
        {
            AppendNonHeaderFooterContent(tmp, node);
        }
    }

    void AppendNonHeaderFooterContent(SharedPtr<Document> tmp, SharedPtr<CompositeNode> node)
    {
        SharedPtr<Node> parentNode = node->get_ParentNode();
        while (!(System::ObjectExt::Is<InlineStory>(parentNode) || System::ObjectExt::Is<Story>(parentNode) ||
                 System::ObjectExt::Is<ShapeBase>(parentNode)))
        {
            auto parent = System::ExplicitCast<CompositeNode>(parentNode->Clone(false));
            parent->AppendChild(node->Clone(true));
            node = parent;

            parentNode = parentNode->get_ParentNode();
        }

        tmp->get_FirstSection()->get_Body()->AppendChild(tmp->ImportNode(node, true, ImportFormatMode::UseDestinationStyles));
    }

    /// <summary>
    /// Adjusts the layout of the document to fit the content area.
    /// </summary>
    void AdjustDocumentLayout(SharedPtr<Document> tmp)
    {
        auto enumerator = MakeObject<LayoutEnumerator>(tmp);
        System::Drawing::RectangleF rect = System::Drawing::RectangleF::Empty;
        rect = CalculateVisibleRect(enumerator, rect);

        tmp->get_FirstSection()->get_PageSetup()->set_PageHeight(rect.get_Height());
        tmp->UpdatePageLayout();
    }

    /// <summary>
    /// Calculates the visible area of the content.
    /// </summary>
    System::Drawing::RectangleF CalculateVisibleRect(SharedPtr<LayoutEnumerator> enumerator, System::Drawing::RectangleF rect)
    {
        System::Drawing::RectangleF result = rect;
        do
        {
            if (enumerator->MoveFirstChild())
            {
                if (enumerator->get_Type() == LayoutEntityType::Line || enumerator->get_Type() == LayoutEntityType::Span)
                {
                    result = result.get_IsEmpty() ? enumerator->get_Rectangle()
                                                  : System::Drawing::RectangleF::Union(result, enumerator->get_Rectangle());
                }
                result = CalculateVisibleRect(enumerator, result);
                enumerator->MoveParent();
            }
        } while (enumerator->MoveNext());

        return result;
    }
    /// <summary>
    /// Finds the minimum bounding box around non-transparent pixels in a Bitmap.
    /// </summary>
    System::Drawing::Rectangle FindBoundingBoxAroundNode(SharedPtr<System::Drawing::Bitmap> originalBitmap)
    {
        System::Drawing::Point min(std::numeric_limits<int>::max(), std::numeric_limits<int>::max());
        System::Drawing::Point max(std::numeric_limits<int>::lowest(), std::numeric_limits<int>::lowest());

        for (int x = 0; x < originalBitmap->get_Width(); ++x)
        {
            for (int y = 0; y < originalBitmap->get_Height(); ++y)
            {
                // Note that you can speed up this part of the algorithm using LockBits and unsafe code instead of GetPixel.
                System::Drawing::Color pixelColor = originalBitmap->GetPixel(x, y);

                // For each pixel that is not transparent, calculate the bounding box around it.
                if (pixelColor.ToArgb() != System::Drawing::Color::Empty.ToArgb())
                {
                    min.set_X(System::Math::Min(x, min.get_X()));
                    min.set_Y(System::Math::Min(y, min.get_Y()));
                    max.set_X(System::Math::Max(x, max.get_X()));
                    max.set_Y(System::Math::Max(y, max.get_Y()));
                }
            }
        }

        // Add one pixel to the width and height to avoid clipping.
        return System::Drawing::Rectangle(min.get_X(), min.get_Y(), max.get_X() - min.get_X() + 1, max.get_Y() - min.get_Y() + 1);
    }
};

}} // namespace DocsExamples::Rendering_and_Printing
