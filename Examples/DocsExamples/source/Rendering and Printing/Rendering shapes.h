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
#include <Aspose.Words.Cpp/FileFormatUtil.h>
#include <Aspose.Words.Cpp/InlineStory.h>
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
#include <Aspose.Words.Cpp/Story.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <drawing/bitmap.h>
#include <drawing/color.h>
#include <drawing/graphics.h>
#include <drawing/graphics_unit.h>
#include <drawing/imaging/image_format.h>
#include <drawing/point.h>
#include <drawing/rectangle.h>
#include <drawing/size.h>
#include <system/details/dispose_guard.h>
#include <system/exceptions.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/memory_stream.h>
#include <system/io/path.h>
#include <system/math.h>
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
        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:RenderShapeAsEmf
        SharedPtr<ShapeRenderer> render = shape->GetShapeRenderer();

        auto imageOptions = MakeObject<ImageSaveOptions>(SaveFormat::Emf);
        imageOptions->set_Scale(1.5f);

        render->Save(ArtifactsDir + u"RenderShape.RenderShapeAsEmf.emf", imageOptions);
        //ExEnd:RenderShapeAsEmf
    }

    void RenderShapeAsJpeg()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:RenderShapeAsJpeg
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
    void RenderShapeToGraphics()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

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
        auto cell = System::DynamicCast<Cell>(doc->GetChild(NodeType::Cell, 2, true));
        RenderNode(cell, ArtifactsDir + u"RenderShape.RenderCellToImage.png", nullptr);
        //ExEnd:RenderCellToImage
    }

    void RenderRowToImage()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        //ExStart:RenderRowToImage
        auto row = System::DynamicCast<Row>(doc->GetChild(NodeType::Row, 0, true));
        RenderNode(row, ArtifactsDir + u"RenderShape.RenderRowToImage.png", nullptr);
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

        RenderNode(textBoxShape->get_LastParagraph(), ArtifactsDir + u"RenderShape.RenderParagraphToImage.png", options);
        //ExEnd:RenderParagraphToImage
    }

    void FindShapeSizes()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:FindShapeSizes
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

        auto shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        //ExStart:RenderShapeImage
        shape->GetShapeRenderer()->Save(ArtifactsDir + u"RenderShape.RenderShapeImage.jpg", nullptr);
        //ExEnd:RenderShapeImage
    }

    /// <summary>
    /// Renders any node in a document to the path specified using the image save options.
    /// </summary>
    /// <param name="node">The node to render.</param>
    /// <param name="filePath">The path to save the rendered image to.</param>
    /// <param name="imageOptions">The image options to use during rendering. This can be null.</param>
    void RenderNode(SharedPtr<Node> node, String filePath, SharedPtr<ImageSaveOptions> imageOptions)
    {
        if (imageOptions == nullptr)
        {
            imageOptions = MakeObject<ImageSaveOptions>(FileFormatUtil::ExtensionToSaveFormat(System::IO::Path::GetExtension(filePath)));
        }

        // Store the paper color to be used on the final image and change to transparent.
        // This will cause any content around the rendered node to be removed later on.
        System::Drawing::Color savePaperColor = imageOptions->get_PaperColor();
        imageOptions->set_PaperColor(System::Drawing::Color::get_Transparent());

        // There a bug which affects the cache of a cloned node.
        // To avoid this, we clone the entire document, including all nodes,
        // finding the matching node in the cloned document and rendering that instead.
        auto doc = System::DynamicCast<Document>(node->get_Document()->Clone(true));
        node = doc->GetChild(NodeType::Any, node->get_Document()->GetChildNodes(NodeType::Any, true)->IndexOf(node), true);

        // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
        // the rendered content of the node.
        auto shape = MakeObject<Shape>(doc, ShapeType::TextBox);
        auto parentSection = System::DynamicCast<Section>(node->GetAncestor(NodeType::Section));

        // Assume that the node cannot be larger than the page in size.
        shape->set_Width(parentSection->get_PageSetup()->get_PageWidth());
        shape->set_Height(parentSection->get_PageSetup()->get_PageHeight());
        shape->set_FillColor(System::Drawing::Color::get_Transparent());

        // Don't draw a surronding line on the shape.
        shape->set_Stroked(false);

        // Move up through the DOM until we find a suitable node to insert into a Shape
        // (a node with a parent can contain paragraphs, tables the same as a shape). Each parent node is cloned
        // on the way up so even a descendant node passed to this method can be rendered. Since we are working
        // with the actual nodes of the document we need to clone the target node into the temporary shape.
        SharedPtr<Node> currentNode = node;
        while (!(System::ObjectExt::Is<InlineStory>(currentNode->get_ParentNode()) || System::ObjectExt::Is<Story>(currentNode->get_ParentNode()) ||
                 System::ObjectExt::Is<ShapeBase>(currentNode->get_ParentNode())))
        {
            auto parent = System::DynamicCast<CompositeNode>(currentNode->get_ParentNode()->Clone(false));
            currentNode = currentNode->get_ParentNode();
            parent->AppendChild(node->Clone(true));
            node = parent;
            // Store this new node to be inserted into the shape.
        }

        // We must add the shape to the document tree to have it rendered.
        shape->AppendChild(node->Clone(true));
        parentSection->get_Body()->get_FirstParagraph()->AppendChild(shape);

        // Render the shape to stream so we can take advantage of the effects of the ImageSaveOptions class.
        // Retrieve the rendered image and remove the shape from the document.
        auto stream = MakeObject<System::IO::MemoryStream>();
        SharedPtr<ShapeRenderer> renderer = shape->GetShapeRenderer();
        renderer->Save(stream, imageOptions);
        shape->Remove();

        System::Drawing::Rectangle crop =
            renderer->GetOpaqueBoundsInPixels(imageOptions->get_Scale(), imageOptions->get_HorizontalResolution(), imageOptions->get_VerticalResolution());

        {
            auto renderedImage = MakeObject<System::Drawing::Bitmap>(stream);
            auto croppedImage = MakeObject<System::Drawing::Bitmap>(crop.get_Width(), crop.get_Height());
            croppedImage->SetResolution(imageOptions->get_HorizontalResolution(), imageOptions->get_VerticalResolution());

            // Create the final image with the proper background color.
            {
                SharedPtr<System::Drawing::Graphics> g = System::Drawing::Graphics::FromImage(croppedImage);
                g->Clear(savePaperColor);
                g->DrawImage(renderedImage, System::Drawing::Rectangle(0, 0, croppedImage->get_Width(), croppedImage->get_Height()), crop.get_X(), crop.get_Y(),
                             crop.get_Width(), crop.get_Height(), System::Drawing::GraphicsUnit::Pixel);

                croppedImage->Save(filePath);
            }
        }
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
