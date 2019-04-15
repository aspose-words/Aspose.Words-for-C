#include "stdafx.h"
#include "examples.h"

#include <drawing/bitmap.h>
#include <Model/Document/Document.h>
#include <Model/Document/FileFormatUtil.h>
#include <Model/Drawing/Shape.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Saving/ImageColorMode.h>
#include <Model/Saving/ImageSaveOptions.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Story.h>
#include <Model/Tables/Cell.h>
#include <Model/Tables/Row.h>
#include <Model/Text/InlineStory.h>
#include <Model/Text/Paragraph.h>
#include <Rendering/ShapeRenderer.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;

namespace
{
    void RenderNode(System::SharedPtr<Node> node, System::String const &filePath, System::SharedPtr<ImageSaveOptions> imageOptions)
    {
        // This code is taken from public API samples of AW.
        // Previously to find opaque bounds of the shape the function
        // That checks every pixel of the rendered image was used.
        // For now opaque bounds is got using ShapeRenderer.GetOpaqueRectangleInPixels method.

        // If no image options are supplied, create default options.
        if (imageOptions == nullptr)
        {
            imageOptions = System::MakeObject<ImageSaveOptions>(FileFormatUtil::ExtensionToSaveFormat(System::IO::Path::GetExtension(filePath)));
        }

        // Store the paper color to be used on the final image and change to transparent.
        // This will cause any content around the rendered node to be removed later on.
        System::Drawing::Color savePaperColor = imageOptions->get_PaperColor();
        imageOptions->set_PaperColor(System::Drawing::Color::get_Transparent());

        // There a bug which affects the cache of a cloned node. To avoid this we instead clone the entire document including all nodes,
        // Find the matching node in the cloned document and render that instead.
        // TODO (std_string) : fix using of overloaded members defined in the base classes
        System::SharedPtr<Document> doc = System::DynamicCast<Document>((System::StaticCast<Node>(node->get_Document()))->Clone(true));
        node = doc->GetChild(NodeType::Any, node->get_Document()->GetChildNodes(NodeType::Any, true)->IndexOf(node), true);

        // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
        // The rendered content of the node.
        System::SharedPtr<Shape> shape = System::MakeObject<Shape>(doc, ShapeType::TextBox);
        System::SharedPtr<Section> parentSection = System::DynamicCast<Section>(node->GetAncestor(NodeType::Section));

        // Assume that the node cannot be larger than the page in size.
        shape->set_Width(parentSection->get_PageSetup()->get_PageWidth());
        shape->set_Height(parentSection->get_PageSetup()->get_PageHeight());
        shape->set_FillColor(System::Drawing::Color::get_Transparent());
        // We must make the shape and paper color transparent.

        // Don't draw a surronding line on the shape.
        shape->set_Stroked(false);

        // Move up through the DOM until we find node which is suitable to insert into a Shape (a node with a parent can contain paragraph, tables the same as a shape).
        // Each parent node is cloned on the way up so even a descendant node passed to this method can be rendered.
        // Since we are working with the actual nodes of the document we need to clone the target node into the temporary shape.
        System::SharedPtr<Node> currentNode = node;
        while (!(System::ObjectExt::Is<InlineStory>(currentNode->get_ParentNode()) || System::ObjectExt::Is<Story>(currentNode->get_ParentNode()) || System::ObjectExt::Is<ShapeBase>(currentNode->get_ParentNode())))
        {
            // TODO (std_string) : fix using of overloaded members defined in the base classes
            System::SharedPtr<CompositeNode> parent = System::DynamicCast<CompositeNode>((System::StaticCast<Node>(currentNode->get_ParentNode()))->Clone(false));
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
        System::SharedPtr<System::IO::MemoryStream> stream = System::MakeObject<System::IO::MemoryStream>();
        System::SharedPtr<ShapeRenderer> renderer = shape->GetShapeRenderer();
        renderer->Save(stream, imageOptions);
        shape->Remove();

        System::Drawing::Rectangle crop = renderer->GetOpaqueBoundsInPixels(imageOptions->get_Scale(), imageOptions->get_HorizontalResolution(), imageOptions->get_VerticalResolution());

        {
            // Load the image into a new bitmap.
            // Clearing resources under 'using' statement
            DisposableHolder<System::Drawing::Bitmap> renderedImageHolder(System::MakeObject<System::Drawing::Bitmap>(stream));
            System::SharedPtr<System::Drawing::Bitmap> croppedImage = System::MakeObject<System::Drawing::Bitmap>(crop.get_Width(), crop.get_Height());
            croppedImage->SetResolution(imageOptions->get_HorizontalResolution(), imageOptions->get_VerticalResolution());

            {
                // Create the final image with the proper background color.
                // Clearing resources under 'using' statement
                DisposableHolder<System::Drawing::Graphics> gHolder(System::Drawing::Graphics::FromImage(croppedImage));
                gHolder.GetObject()->Clear(savePaperColor);
                gHolder.GetObject()->DrawImage(renderedImageHolder.GetObject(),
                                     System::Drawing::Rectangle(0, 0, croppedImage->get_Width(), croppedImage->get_Height()),
                                     crop.get_X(),
                                     crop.get_Y(),
                                     crop.get_Width(),
                                     crop.get_Height(),
                                     System::Drawing::GraphicsUnit::Pixel);
                croppedImage->Save(filePath);
            }
        }
    }

    void RenderShapeToDisk(System::String const &dataDir, System::SharedPtr<Shape> shape)
    {
        // ExStart:RenderShapeToDisk
        System::SharedPtr<ShapeRenderer> r = shape->GetShapeRenderer();

        // Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
        System::SharedPtr<ImageSaveOptions> imageOptions = System::MakeObject<ImageSaveOptions>(SaveFormat::Emf);
        imageOptions->set_Scale(1.5f);

        System::String outputPath = dataDir + GetOutputFilePath(u"RenderShape.RenderShapeToDisk.emf");
        // Save the rendered image to disk.
        r->Save(outputPath, imageOptions);
        // ExEnd:RenderShapeToDisk
        std::cout << "Shape rendered to disk successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void RenderShapeToStream(System::String const &dataDir, System::SharedPtr<Shape> shape)
    {
        // ExStart:RenderShapeToStream
        System::SharedPtr<ShapeRenderer> r = System::MakeObject<ShapeRenderer>(shape);

        // Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
        System::SharedPtr<ImageSaveOptions> imageOptions = System::MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);
        imageOptions->set_ImageColorMode(ImageColorMode::Grayscale);
        imageOptions->set_ImageBrightness(0.45f);
        System::String outputPath = dataDir + GetOutputFilePath(u"RenderShape.RenderShapeToStream.jpg");
        System::SharedPtr<System::IO::FileStream> stream = System::MakeObject<System::IO::FileStream>(outputPath, System::IO::FileMode::Create);

        // Save the rendered image to the stream using different options.
        r->Save(stream, imageOptions);
        // ExEnd:RenderShapeToStream
        std::cout << "Shape rendered to stream successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void RenderShapeToGraphics(System::String const &dataDir, System::SharedPtr<Shape> shape)
    {
        System::SharedPtr<ShapeRenderer> r = shape->GetShapeRenderer();

        // Find the size that the shape will be rendered to at the specified scale and resolution.
        System::Drawing::Size shapeSizeInPixels = r->GetSizeInPixels(1.0f, 96.0f);

        // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
        // And make sure that the graphics canvas is large enough to compensate for this.
        int32_t maxSide = System::Math::Max(shapeSizeInPixels.get_Width(), shapeSizeInPixels.get_Height());

        {
            // Clearing resources under 'using' statement
            DisposableHolder<System::Drawing::Bitmap> imageHolder(System::MakeObject<System::Drawing::Bitmap>((int32_t)(maxSide * 1.25), (int32_t)(maxSide * 1.25)));

            {
                // Rendering to a graphics object means we can specify settings and transformations to be applied to 
                // The shape that is rendered. In our case we will rotate the rendered shape.
                // Clearing resources under 'using' statement
                DisposableHolder<System::Drawing::Graphics> grHolder(System::Drawing::Graphics::FromImage(imageHolder.GetObject()));

                // Clear the shape with the background color of the document.
                grHolder.GetObject()->Clear(shape->get_Document()->get_PageColor());
                // Center the rotation using translation method below
                grHolder.GetObject()->TranslateTransform((float)imageHolder.GetObject()->get_Width() / 8, (float)imageHolder.GetObject()->get_Height() / 2);
                // Rotate the image by 45 degrees.
                grHolder.GetObject()->RotateTransform(45.0f);
                // Undo the translation.
                grHolder.GetObject()->TranslateTransform(-(float)imageHolder.GetObject()->get_Width() / 8, -(float)imageHolder.GetObject()->get_Height() / 2);

                // Render the shape onto the graphics object.
                r->RenderToSize(grHolder.GetObject(), 0.0f, 0.0f, shapeSizeInPixels.get_Width(), shapeSizeInPixels.get_Height());
            }

            System::String outputPath = dataDir + GetOutputFilePath(u"RenderShape.RenderShapeToGraphics.png");
            imageHolder.GetObject()->Save(outputPath, System::Drawing::Imaging::ImageFormat::get_Png());
            std::cout << "Shape rendered to graphics successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
        }
    }

    void RenderCellToImage(System::String const &dataDir, System::SharedPtr<Document> doc)
    {
        // ExStart:RenderCellToImage
        System::SharedPtr<Cell> cell = System::DynamicCast<Cell>(doc->GetChild(NodeType::Cell, 2, true));
        // The third cell in the first table.
        System::String outputPath = dataDir + GetOutputFilePath(u"RenderShape.RenderCellToImage.png");
        RenderNode(cell, outputPath, nullptr);
        // ExEnd:RenderCellToImage
        std::cout << "Cell rendered to image successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void RenderRowToImage(System::String const &dataDir, System::SharedPtr<Document> doc)
    {
        // ExStart:RenderRowToImage
        System::SharedPtr<Row> row = System::DynamicCast<Row>(doc->GetChild(NodeType::Row, 0, true));
        // The first row in the first table.

        System::String outputPath = dataDir + GetOutputFilePath(u"RenderShape.RenderRowToImage.png");
        RenderNode(row, outputPath, nullptr);
        // ExEnd:RenderRowToImage
        std::cout << "Row rendered to image successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void RenderParagraphToImage(System::String const &dataDir, System::SharedPtr<Document> doc)
    {
        // ExStart:RenderParagraphToImage
        System::SharedPtr<Shape> shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        System::SharedPtr<Paragraph> paragraph = shape->get_LastParagraph();

        // Save the node with a light pink background.
        System::SharedPtr<ImageSaveOptions> options = System::MakeObject<ImageSaveOptions>(SaveFormat::Png);
        options->set_PaperColor(System::Drawing::Color::get_LightPink());
        System::String outputPath = dataDir + GetOutputFilePath(u"RenderShape.RenderParagraphToImage.png");
        RenderNode(paragraph, outputPath, options);
        // ExEnd:RenderParagraphToImage
        std::cout << "Paragraph rendered to image successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void FindShapeSizes(System::SharedPtr<Shape> shape)
    {
        // ExStart:FindShapeSizes
        System::Drawing::SizeF shapeSizeInDocument = shape->GetShapeRenderer()->get_SizeInPoints();
        float width = shapeSizeInDocument.get_Width();
        // The width of the shape.
        float height = shapeSizeInDocument.get_Height();
        // The height of the shape.

        System::Drawing::Size shapeRenderedSize = shape->GetShapeRenderer()->GetSizeInPixels(1.0f, 96.0f);

        {
            // Clearing resources under 'using' statement
            DisposableHolder<System::Drawing::Bitmap> imageHolder(System::MakeObject<System::Drawing::Bitmap>(shapeRenderedSize.get_Width(), shapeRenderedSize.get_Height()));
            {
                // Clearing resources under 'using' statement
                DisposableHolder<System::Drawing::Graphics> gHolder(System::Drawing::Graphics::FromImage(imageHolder.GetObject()));
                // Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
            }
        }
        // ExEnd:FindShapeSizes
    }

    void RenderShapeImage(System::String const &dataDir, System::SharedPtr<Shape> shape)
    {
        // ExStart:RenderShapeImage
        System::String outputPath = dataDir + GetOutputFilePath(u"RenderShape.RenderShapeImage.jpg");
        // Save the Shape image to disk in JPEG format and using default options.
        shape->GetShapeRenderer()->Save(outputPath, nullptr);
        // ExEnd:RenderShapeImage
        std::cout << "Shape image rendered successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void RenderShape()
{
    std::cout << "RenderShape example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_RenderingAndPrinting();

    // Load the documents which store the shapes we want to render.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile RenderShape.doc");
    System::SharedPtr<Document> doc2 = System::MakeObject<Document>(dataDir + u"TestFile RenderShape.docx");

    // Retrieve the target shape from the document. In our sample document this is the first shape.
    System::SharedPtr<Shape> shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

    // Test rendering of different types of nodes.
    RenderShapeToDisk(dataDir, shape);
    RenderShapeToStream(dataDir, shape);
    RenderShapeToGraphics(dataDir, shape);
    RenderCellToImage(dataDir, doc);
    RenderRowToImage(dataDir, doc);
    RenderParagraphToImage(dataDir, doc);
    FindShapeSizes(shape);
    RenderShapeImage(dataDir, shape);
    std::cout << "RenderShape example finished." << std::endl << std::endl;
}