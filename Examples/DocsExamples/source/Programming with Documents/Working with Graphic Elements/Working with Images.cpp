#include "Programming with Documents/Working with Graphic Elements/Working with Images.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <drawing/bitmap.h>
#include <drawing/drawing2d/interpolation_mode.h>
#include <drawing/graphics.h>
#include <drawing/imaging/encoder.h>
#include <drawing/imaging/encoder_parameter.h>
#include <drawing/imaging/encoder_parameters.h>
#include <system/array.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/guid.h>
#include <system/io/memory_stream.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Layout;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements {

int32_t Resampler::Resample(System::SharedPtr<Document> doc, int32_t desiredPpi, int32_t jpegQuality)
{
    int32_t count = 0;

    for (const auto& shape : System::IterateOver<Shape>(doc->GetChildNodes(NodeType::Shape, true)))
    {
        // It is important to use this method to get the picture shape size in points correctly,
        // even if it is inside a group shape.
        System::Drawing::SizeF shapeSizeInPoints = shape->get_SizeInPoints();

        if (ResampleCore(shape->get_ImageData(), shapeSizeInPoints, desiredPpi, jpegQuality))
        {
            count++;
        }
    }

    return count;
}

bool Resampler::ResampleCore(System::SharedPtr<ImageData> imageData, System::Drawing::SizeF shapeSizeInPoints, int32_t ppi, int32_t jpegQuality)
{
    // The are several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
    if (imageData == nullptr)
    {
        return false;
    }

    // An image can be stored in shape or linked somewhere else, let's skip images that do not store bytes in shape.
    System::ArrayPtr<uint8_t> originalBytes = imageData->get_ImageBytes();
    if (originalBytes == nullptr)
    {
        return false;
    }

    // Ignore metafiles, they are vector drawings, and we don't want to resample them.
    ImageType imageType = imageData->get_ImageType();
    if (imageType == ImageType::Wmf || imageType == ImageType::Emf)
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

        std::cout << "Image PpiX:" << ((int32_t)currentPpiX) << ", PpiY:" << ((int32_t)currentPpiY) << ". ";

        // Let's resample only if the current PPI is higher than the requested PPI (e.g., we have extra data we can get rid of).
        if (currentPpiX <= ppi || currentPpiY <= ppi)
        {
            std::cout << "Skipping." << std::endl;
            return false;
        }

        {
            System::SharedPtr<System::Drawing::Image> srcImage = imageData->ToImage();
            // Create a new image of such size that it will hold only the pixels required by the desired PPI.
            int32_t dstWidthPixels = (int32_t)(shapeWidthInches * ppi);
            int32_t dstHeightPixels = (int32_t)(shapeHeightInches * ppi);
            {
                auto dstImage = System::MakeObject<System::Drawing::Bitmap>(dstWidthPixels, dstHeightPixels);
                // Drawing the source image to the new image scales it to the new size.
                {
                    System::SharedPtr<System::Drawing::Graphics> gr = System::Drawing::Graphics::FromImage(dstImage);
                    gr->set_InterpolationMode(System::Drawing::Drawing2D::InterpolationMode::HighQualityBicubic);
                    gr->DrawImage(srcImage, 0, 0, dstWidthPixels, dstHeightPixels);
                }

                // Create JPEG encoder parameters with the quality setting.
                System::SharedPtr<System::Drawing::Imaging::ImageCodecInfo> encoderInfo = GetEncoderInfo(System::Drawing::Imaging::ImageFormat::get_Jpeg());
                auto encoderParams = System::MakeObject<System::Drawing::Imaging::EncoderParameters>();
                encoderParams->get_Param()[0] = System::MakeObject<System::Drawing::Imaging::EncoderParameter>(System::Drawing::Imaging::Encoder::Quality,
                                                                                                               static_cast<int64_t>(jpegQuality));

                // Save the image as JPEG to a memory stream.
                auto dstStream = System::MakeObject<System::IO::MemoryStream>();
                dstImage->Save(dstStream, encoderInfo, encoderParams);

                // If the image saved as JPEG is smaller than the original, store it in shape.
                std::cout << "Original size " << originalBytes->get_Length() << ", new size " << dstStream->get_Length() << "." << std::endl;
                if (dstStream->get_Length() < originalBytes->get_Length())
                {
                    dstStream->set_Position(0);
                    imageData->SetImage(dstStream);
                    return true;
                }
            }
        }
    }
    catch (System::Exception& e)
    {
        // Catch an exception, log an error, and continue to process one of the images for whatever reason.
        std::cout << (System::String(u"Error processing an image, ignoring. ") + e->get_Message()) << std::endl;
    }

    return false;
}

System::SharedPtr<System::Drawing::Imaging::ImageCodecInfo> Resampler::GetEncoderInfo(System::SharedPtr<System::Drawing::Imaging::ImageFormat> format)
{
    System::ArrayPtr<System::SharedPtr<System::Drawing::Imaging::ImageCodecInfo>> encoders = System::Drawing::Imaging::ImageCodecInfo::GetImageEncoders();

    for (System::SharedPtr<System::Drawing::Imaging::ImageCodecInfo> codecInfo : encoders)
    {
        if (codecInfo->get_FormatID() == format->get_Guid())
        {
            return codecInfo;
        }
    }

    throw System::Exception(u"Cannot find a codec.");
}

namespace gtest_test {

class WorkingWithImages : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithImages> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithImages>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithImages> WorkingWithImages::s_instance;

TEST_F(WorkingWithImages, AddImageToEachPage)
{
    s_instance->AddImageToEachPage();
}

TEST_F(WorkingWithImages, InsertBarcodeImage)
{
    s_instance->InsertBarcodeImage();
}

TEST_F(WorkingWithImages, CompressImages)
{
    s_instance->CompressImages();
}

} // namespace gtest_test

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements
