#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/special_casts.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/guid.h>
#include <system/exceptions.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/Node.h>
#include <Model/Drawing/Shape.h>
#include <Model/Drawing/ImageType.h>
#include <Model/Drawing/ImageSize.h>
#include <Model/Drawing/ImageData.h>
#include <Model/Document/Document.h>
#include <Model/Document/ConvertUtil.h>
#include <drawing/size_f.h>
#include <drawing/imaging/image_format.h>
#include <drawing/imaging/image_codec_info.h>
#include <drawing/imaging/encoder_parameters.h>
#include <drawing/imaging/encoder_parameter.h>
#include <drawing/imaging/encoder.h>
#include <drawing/image.h>
#include <drawing/graphics.h>
#include <drawing/drawing2d/interpolation_mode.h>
#include <drawing/bitmap.h>
#include <cstdint>
#include <system/diagnostics/debug.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;

namespace
{
    int32_t GetFileSize(const System::String& fileName)
    {
        auto stream = System::IO::File::OpenRead(fileName);
        return stream->get_Length();
    }

    System::SharedPtr<System::Drawing::Imaging::ImageCodecInfo> GetEncoderInfo(const System::SharedPtr<System::Drawing::Imaging::ImageFormat>& format)
    {
        auto encoders = System::Drawing::Imaging::ImageCodecInfo::GetImageEncoders();

        for (auto& encoder : encoders->data())
        {
            if (encoder->get_FormatID() == format->get_Guid())
            {
                return encoder;
            }
        }
        throw System::Exception(u"Cannot find a codec.");
    }

    bool ResampleCore(const System::SharedPtr<ImageData>& imageData, System::Drawing::SizeF shapeSizeInPoints, int32_t ppi, int32_t jpegQuality)
    {
        // The are actually several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
        if (imageData == nullptr)
        {
            return false;
        }
        // An image can be stored in the shape or linked from somewhere else. Let's skip images that do not store bytes in the shape.
        System::ArrayPtr<uint8_t> originalBytes = imageData->get_ImageBytes();
        if (originalBytes == nullptr)
        {
            return false;
        }

        // Ignore metafiles, they are vector drawings and we don't want to resample them.
        ImageType imageType = imageData->get_ImageType();
        if (System::ObjectExt::Equals(imageType, Aspose::Words::Drawing::ImageType::Wmf) || System::ObjectExt::Equals(imageType, Aspose::Words::Drawing::ImageType::Emf))
        {
            return false;
        }

        try
        {
            double shapeWidthInches = ConvertUtil::PointToInch(shapeSizeInPoints.get_Width());
            double shapeHeightInches = ConvertUtil::PointToInch(shapeSizeInPoints.get_Height());

            // Calculate the current PPI of the image.
            System::SharedPtr<ImageSize> imageSize = imageData->get_ImageSize();
            double currentPpiX = imageSize->get_WidthPixels() / shapeWidthInches;
            double currentPpiY = imageSize->get_HeightPixels() / shapeHeightInches;
            std::cout << "Image PpiX:" << currentPpiX << ", PpiY:" << currentPpiY << ". ";

            // Let's resample only if the current PPI is higher than the requested PPI (e.g. we have extra data we can get rid of).
            if ((currentPpiX <= ppi) || (currentPpiY <= ppi))
            {
                std::cout << "Skipping.";
                return false;
            }

            System::SharedPtr<System::Drawing::Image> srcImage = imageData->ToImage();
            // Create a new image of such size that it will hold only the pixels required by the desired ppi.
            int32_t dstWidthPixels = (int32_t)(shapeWidthInches * ppi);
            int32_t dstHeightPixels = (int32_t)(shapeHeightInches * ppi);
            System::SharedPtr<System::Drawing::Bitmap> dstImage = System::MakeObject<System::Drawing::Bitmap>(dstWidthPixels, dstHeightPixels);
            // Drawing the source image to the new image scales it to the new size.
            System::SharedPtr<System::Drawing::Graphics> gr = System::Drawing::Graphics::FromImage(dstImage);
            gr->set_InterpolationMode(System::Drawing::Drawing2D::InterpolationMode::HighQualityBicubic);
            gr->DrawImage(srcImage, 0, 0, dstWidthPixels, dstHeightPixels);
            // Create JPEG encoder parameters with the quality setting.
            System::SharedPtr<System::Drawing::Imaging::ImageCodecInfo> encoderInfo = GetEncoderInfo(System::Drawing::Imaging::ImageFormat::get_Jpeg());
            System::SharedPtr<System::Drawing::Imaging::EncoderParameters> encoderParams = System::MakeObject<System::Drawing::Imaging::EncoderParameters>();
            encoderParams->get_Param()[0] = System::MakeObject<System::Drawing::Imaging::EncoderParameter>(System::Drawing::Imaging::Encoder::Quality, jpegQuality);

            // Save the image as JPEG to a memory stream.
            System::SharedPtr<System::IO::MemoryStream> dstStream = System::MakeObject<System::IO::MemoryStream>();
            dstImage->Save(dstStream, encoderInfo, encoderParams);

            // If the image saved as JPEG is smaller than the original, store it in the shape.
            std::cout << "Original size " << originalBytes->get_Length() << ", new size " << dstStream->get_Length() << "." << std::endl;
            if (dstStream->get_Length() < originalBytes->get_Length())
            {
                dstStream->set_Position(0);
                imageData->SetImage(dstStream);
                return true;
            }
        }
        catch (System::Exception& e)
        {
            // Catch an exception, log an error and continue if cannot process one of the images for whatever reason.
            std::cout << "Error processing an image, ignoring. " << e.get_Message().ToUtf8String() << std::endl;
        }
        return false;
    }

    int32_t Resample(const System::SharedPtr<Document>& doc, int32_t desiredPpi, int32_t jpegQuality)
    {
        int32_t count = 0;

        // Convert VML shapes.
        auto shape_enumerator = (System::DynamicCastEnumerableTo<System::SharedPtr<Shape>>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)))->GetEnumerator();
        decltype(shape_enumerator->get_Current()) shape;
        while (shape_enumerator->MoveNext() && (shape = shape_enumerator->get_Current(), true))
        {
            // It is important to use this method to correctly get the picture shape size in points even if the picture is inside a group shape.
            System::Drawing::SizeF shapeSizeInPoints = shape->get_SizeInPoints();

            if (ResampleCore(shape->get_ImageData(), shapeSizeInPoints, desiredPpi, jpegQuality))
            {
                count++;
            }
        }
        return count;
    }
}


void CompressImages()
{
    std::cout << "CompressImages example started." << std::endl;
    // ExStart:CompressImages
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithImages();
    System::String srcFileName = dataDir + u"Test.doc";

    std::cout << "Loading " << srcFileName.ToUtf8String() << ". Size " << GetFileSize(srcFileName) << "." << std::endl;
    System::SharedPtr<Document> doc = System::MakeObject<Document>(srcFileName);

    // 220ppi Print - said to be excellent on most printers and screens.
    // 150ppi Screen - said to be good for web pages and projectors.
    // 96ppi Email - said to be good for minimal document size and sharing.
    const int32_t desiredPpi = 150;
    
    // In .NET this seems to be a good compression / quality setting.
    const int32_t jpegQuality = 90;
    
    // Resample images to desired ppi and save.
    int32_t count = Resample(doc, desiredPpi, jpegQuality);
    
    std::cout << "Resampled " << count << " images." << std::endl;
    
    if (count != 1)
    {
        std::cout << "We expected to have only 1 image resampled in this test document!" << std::endl;
    }
    
    System::String dstFileName = dataDir + GetOutputFilePath(u"CompressImages.doc");
    doc->Save(dstFileName);
    std::cout << "Saving " << dstFileName.ToUtf8String() << ". Size " << GetFileSize(dstFileName) << "." << std::endl;
    
    // Verify that the first image was compressed by checking the new Ppi.
    doc = System::MakeObject<Document>(dstFileName);
    System::SharedPtr<Shape> shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
    double imagePpi = shape->get_ImageData()->get_ImageSize()->get_WidthPixels() / ConvertUtil::PointToInch(shape->get_SizeInPoints().get_Width());
    assert(imagePpi < 150, u"Image was not resampled successfully.");
    // ExEnd:CompressImages
    std::cout << "Compressed images successfully." << std::endl << "File saved at " << dstFileName.ToUtf8String() << std::endl;
    std::cout << "CompressImages example finished." << std::endl << std::endl;
}


