#pragma once

#ifdef ASPOSE_BARCODE_AVAILABLE

#include <cstdint>
#include <Aspose.Words.Cpp/Fields/BarcodeParameters.h>
#include <Aspose.Words.Cpp/Fields/IBarcodeGenerator.h>
// Aspose.BarCode for C++ headers. Its NuGet package puts "include" on the
// include path, so these are referenced relative to that root.
#include <BarCode.Generation/BarcodeGenerator.h>
#include <BarCode.Generation/EncodeTypes/BaseEncodeType.h>
#include <BarCode.Generation/EncodeTypes/EncodeTypes.h>
#include <BarCode.Generation/EncodeTypes/SymbologyEncodeType.h>
#include <BarCode.Generation/GenerationParameters/BarcodeParameters.h>
#include <BarCode.Generation/GenerationParameters/BaseGenerationParameters.h>
#include <BarCode.Generation/GenerationParameters/BorderDashStyle.h>
#include <BarCode.Generation/GenerationParameters/BorderParameters.h>
#include <BarCode.Generation/GenerationParameters/CodetextParameters.h>
#include <BarCode.Generation/GenerationParameters/QRErrorLevel.h>
#include <BarCode.Generation/Helpers/Unit.h>
#include <drawing/bitmap.h>
#include <drawing/brushes.h>
#include <drawing/color.h>
#include <drawing/font.h>
#include <drawing/font_style.h>
#include <drawing/graphics.h>
#include <drawing/imaging/image_format.h>
#include <drawing/rectangle.h>
#include <system/convert.h>
#include <system/exceptions.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>
#include <system/math.h>
#include <system/string.h>

using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples {

class CustomBarcodeGeneratorUtils
{
public:
    static constexpr double DefaultQRXDimensionInPixels = 4.0;
    static constexpr double Default1DXDimensionInPixels = 1.0;

    /// <summary>
    /// Converts a height value in twips to pixels using a default DPI of 96.
    /// </summary>
    /// <param name="heightInTwips">The height value in twips.</param>
    /// <param name="defVal">The default value to return if the conversion fails.</param>
    /// <returns>The height value in pixels.</returns>
    static double TwipsToPixels(String heightInTwips, double defVal)
    {
        return TwipsToPixels(heightInTwips, 96, defVal);
    }

    /// <summary>
    /// Converts a height value in twips to pixels based on the given resolution.
    /// </summary>
    /// <param name="heightInTwips">The height value in twips to be converted.</param>
    /// <param name="resolution">The resolution in pixels per inch.</param>
    /// <param name="defVal">The default value to be returned if the conversion fails.</param>
    /// <returns>The converted height value in pixels.</returns>
    static double TwipsToPixels(String heightInTwips, double resolution, double defVal)
    {
        try
        {
            int lVal = System::Convert::ToInt32(heightInTwips);
            return (lVal / 1440.0) * resolution;
        }
        catch (...)
        {
            return defVal;
        }
    }

    /// <summary>
    /// Gets the rotation angle in degrees based on the given rotation angle string.
    /// </summary>
    /// <param name="rotationAngle">The rotation angle string.</param>
    /// <param name="defVal">The default value to return if the rotation angle is not recognized.</param>
    /// <returns>The rotation angle in degrees.</returns>
    static float GetRotationAngle(String rotationAngle, float defVal)
    {
        if (rotationAngle == u"0")
        {
            return 0;
        }
        if (rotationAngle == u"1")
        {
            return 270;
        }
        if (rotationAngle == u"2")
        {
            return 180;
        }
        if (rotationAngle == u"3")
        {
            return 90;
        }
        return defVal;
    }

    /// <summary>
    /// Converts a string representation of an error correction level to a QRErrorLevel enum value.
    /// </summary>
    /// <param name="errorCorrectionLevel">The string representation of the error correction level.</param>
    /// <param name="def">The default error correction level to return if the input is invalid.</param>
    /// <returns>The corresponding QRErrorLevel enum value.</returns>
    static Aspose::BarCode::Generation::QRErrorLevel GetQRCorrectionLevel(String errorCorrectionLevel, Aspose::BarCode::Generation::QRErrorLevel def)
    {
        if (errorCorrectionLevel == u"0")
        {
            return Aspose::BarCode::Generation::QRErrorLevel::LevelL;
        }
        if (errorCorrectionLevel == u"1")
        {
            return Aspose::BarCode::Generation::QRErrorLevel::LevelM;
        }
        if (errorCorrectionLevel == u"2")
        {
            return Aspose::BarCode::Generation::QRErrorLevel::LevelQ;
        }
        if (errorCorrectionLevel == u"3")
        {
            return Aspose::BarCode::Generation::QRErrorLevel::LevelH;
        }
        return def;
    }

    /// <summary>
    /// Gets the barcode encode type based on the given encode type from Word.
    /// </summary>
    /// <param name="encodeTypeFromWord">The encode type from Word.</param>
    /// <returns>The barcode encode type.</returns>
    static SharedPtr<Aspose::BarCode::Generation::SymbologyEncodeType> GetBarcodeEncodeType(String encodeTypeFromWord)
    {
        // https://support.microsoft.com/en-au/office/field-codes-displaybarcode-6d81eade-762d-4b44-ae81-f9d3d9e07be3
        if (encodeTypeFromWord == u"QR")
        {
            return Aspose::BarCode::Generation::EncodeTypes::QR;
        }
        if (encodeTypeFromWord == u"CODE128")
        {
            return Aspose::BarCode::Generation::EncodeTypes::Code128;
        }
        if (encodeTypeFromWord == u"CODE39")
        {
            return Aspose::BarCode::Generation::EncodeTypes::Code39;
        }
        if (encodeTypeFromWord == u"JPPOST")
        {
            return Aspose::BarCode::Generation::EncodeTypes::RM4SCC;
        }
        if (encodeTypeFromWord == u"EAN8" || encodeTypeFromWord == u"JAN8")
        {
            return Aspose::BarCode::Generation::EncodeTypes::EAN8;
        }
        if (encodeTypeFromWord == u"EAN13" || encodeTypeFromWord == u"JAN13")
        {
            return Aspose::BarCode::Generation::EncodeTypes::EAN13;
        }
        if (encodeTypeFromWord == u"UPCA")
        {
            return Aspose::BarCode::Generation::EncodeTypes::UPCA;
        }
        if (encodeTypeFromWord == u"UPCE")
        {
            return Aspose::BarCode::Generation::EncodeTypes::UPCE;
        }
        if (encodeTypeFromWord == u"CASE" || encodeTypeFromWord == u"ITF14")
        {
            return Aspose::BarCode::Generation::EncodeTypes::ITF14;
        }
        if (encodeTypeFromWord == u"NW7")
        {
            return Aspose::BarCode::Generation::EncodeTypes::Codabar;
        }
        return Aspose::BarCode::Generation::EncodeTypes::None;
    }

    /// <summary>
    /// Converts a hexadecimal color string to a Color object.
    /// </summary>
    /// <param name="inputColor">The hexadecimal color string to convert.</param>
    /// <param name="defVal">The default Color value to return if the conversion fails.</param>
    /// <returns>The Color object representing the converted color, or the default value if the conversion fails.</returns>
    static System::Drawing::Color ConvertColor(String inputColor, System::Drawing::Color defVal)
    {
        if (String::IsNullOrEmpty(inputColor))
        {
            return defVal;
        }
        try
        {
            int color = System::Convert::ToInt32(inputColor, 16);
            // Return Color::FromArgb((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
            return System::Drawing::Color::FromArgb(color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF);
        }
        catch (...)
        {
            return defVal;
        }
    }

    /// <summary>
    /// Calculates the scale factor based on the provided string representation.
    /// </summary>
    /// <param name="scaleFactor">The string representation of the scale factor.</param>
    /// <param name="defVal">The default value to return if the scale factor cannot be parsed.</param>
    /// <returns>
    /// The scale factor as a decimal value between 0 and 1, or the default value if the scale factor cannot be parsed.
    /// </returns>
    static double ScaleFactor(String scaleFactor, double defVal)
    {
        try
        {
            int scale = System::Convert::ToInt32(scaleFactor);
            return scale / 100.0;
        }
        catch (...)
        {
            return defVal;
        }
    }

    /// <summary>
    /// Sets the position code style for a barcode generator.
    /// </summary>
    /// <param name="gen">The barcode generator.</param>
    /// <param name="posCodeStyle">The position code style to set.</param>
    /// <param name="barcodeValue">The barcode value.</param>
    static void SetPosCodeStyle(SharedPtr<Aspose::BarCode::Generation::BarcodeGenerator> gen, String posCodeStyle, String barcodeValue)
    {
        // STD default and without changes.
        if (posCodeStyle == u"SUP2")
        {
            gen->set_CodeText(barcodeValue.Substring(0, barcodeValue.get_Length() - 2));
            gen->get_Parameters()->get_Barcode()->get_Supplement()->set_SupplementData(barcodeValue.Substring(barcodeValue.get_Length() - 2, 2));
        }
        else if (posCodeStyle == u"SUP5")
        {
            gen->set_CodeText(barcodeValue.Substring(0, barcodeValue.get_Length() - 5));
            gen->get_Parameters()->get_Barcode()->get_Supplement()->set_SupplementData(barcodeValue.Substring(barcodeValue.get_Length() - 5, 5));
        }
        else if (posCodeStyle == u"CASE")
        {
            gen->get_Parameters()->get_Border()->set_Visible(true);
            gen->get_Parameters()->get_Border()->set_Color(gen->get_Parameters()->get_Barcode()->get_BarColor());
            gen->get_Parameters()->get_Border()->set_DashStyle(Aspose::BarCode::Generation::BorderDashStyle::Solid);
            gen->get_Parameters()->get_Border()->get_Width()->set_Pixels(gen->get_Parameters()->get_Barcode()->get_XDimension()->get_Pixels() * 5);
        }
    }

    /// <summary>
    /// Draws an error image with the specified error message.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <returns>A Bitmap object representing the error image.</returns>
    static SharedPtr<System::Drawing::Bitmap> DrawErrorImage(String message)
    {
        auto bmp = MakeObject<System::Drawing::Bitmap>(100, 100);

        {
            SharedPtr<System::Drawing::Graphics> grf = System::Drawing::Graphics::FromImage(bmp);
            grf->DrawString(message,
                            MakeObject<System::Drawing::Font>(u"Microsoft Sans Serif", 8.0f, System::Drawing::FontStyle::Regular),
                            System::Drawing::Brushes::get_Red(),
                            System::Drawing::Rectangle(0, 0, bmp->get_Width(), bmp->get_Height()));
        }

        return bmp;
    }

    static SharedPtr<System::IO::Stream> ConvertImageToWord(SharedPtr<System::Drawing::Bitmap> bmp)
    {
        auto ms = MakeObject<System::IO::MemoryStream>();
        bmp->Save(ms, System::Drawing::Imaging::ImageFormat::get_Png());
        ms->set_Position(0);

        return ms;
    }
};

class CustomBarcodeGenerator : public Aspose::Words::Fields::IBarcodeGenerator
{
public:
    SharedPtr<System::IO::Stream> GetBarcodeImage(SharedPtr<Aspose::Words::Fields::BarcodeParameters> parameters) override
    {
        try
        {
            auto gen = MakeObject<Aspose::BarCode::Generation::BarcodeGenerator>(
                System::StaticCast<Aspose::BarCode::Generation::BaseEncodeType>(
                    CustomBarcodeGeneratorUtils::GetBarcodeEncodeType(parameters->get_BarcodeType())),
                parameters->get_BarcodeValue());

            // Set color.
            gen->get_Parameters()->get_Barcode()->set_BarColor(
                CustomBarcodeGeneratorUtils::ConvertColor(parameters->get_ForegroundColor(), gen->get_Parameters()->get_Barcode()->get_BarColor()));
            gen->get_Parameters()->set_BackColor(
                CustomBarcodeGeneratorUtils::ConvertColor(parameters->get_BackgroundColor(), gen->get_Parameters()->get_BackColor()));

            // Set display or hide text.
            if (!parameters->get_DisplayText())
            {
                gen->get_Parameters()->get_Barcode()->get_CodeTextParameters()->set_Location(Aspose::BarCode::Generation::CodeLocation::None);
            }
            else
            {
                gen->get_Parameters()->get_Barcode()->get_CodeTextParameters()->set_Location(Aspose::BarCode::Generation::CodeLocation::Below);
            }

            // Set QR Code error correction level.
            gen->get_Parameters()->get_Barcode()->get_QR()->set_ErrorLevel(Aspose::BarCode::Generation::QRErrorLevel::LevelH);
            if (!String::IsNullOrEmpty(parameters->get_ErrorCorrectionLevel()))
            {
                gen->get_Parameters()->get_Barcode()->get_QR()->set_ErrorLevel(CustomBarcodeGeneratorUtils::GetQRCorrectionLevel(
                    parameters->get_ErrorCorrectionLevel(), gen->get_Parameters()->get_Barcode()->get_QR()->get_ErrorLevel()));
            }

            // Set rotation angle.
            if (!String::IsNullOrEmpty(parameters->get_SymbolRotation()))
            {
                gen->get_Parameters()->set_RotationAngle(
                    CustomBarcodeGeneratorUtils::GetRotationAngle(parameters->get_SymbolRotation(), gen->get_Parameters()->get_RotationAngle()));
            }

            // Set scaling factor.
            double scalingFactor = 1;
            if (!String::IsNullOrEmpty(parameters->get_ScalingFactor()))
            {
                scalingFactor = CustomBarcodeGeneratorUtils::ScaleFactor(parameters->get_ScalingFactor(), scalingFactor);
            }

            // Set size.
            if (gen->get_BarcodeType() == System::StaticCast<Aspose::BarCode::Generation::BaseEncodeType>(Aspose::BarCode::Generation::EncodeTypes::QR))
            {
                gen->get_Parameters()->get_Barcode()->get_XDimension()->set_Pixels(
                    (float)System::Math::Max(1.0, System::Math::Round(CustomBarcodeGeneratorUtils::DefaultQRXDimensionInPixels * scalingFactor)));
            }
            else
            {
                gen->get_Parameters()->get_Barcode()->get_XDimension()->set_Pixels(
                    (float)System::Math::Max(1.0, System::Math::Round(CustomBarcodeGeneratorUtils::Default1DXDimensionInPixels * scalingFactor)));
            }

            // Set height.
            if (!String::IsNullOrEmpty(parameters->get_SymbolHeight()))
            {
                gen->get_Parameters()->get_Barcode()->get_BarHeight()->set_Pixels((float)System::Math::Max(
                    5.0,
                    System::Math::Round(CustomBarcodeGeneratorUtils::TwipsToPixels(parameters->get_SymbolHeight(),
                                                                                  gen->get_Parameters()->get_Barcode()->get_BarHeight()->get_Pixels()) *
                                        scalingFactor)));
            }

            // Set style of a Point-of-Sale barcode.
            if (!String::IsNullOrEmpty(parameters->get_PosCodeStyle()))
            {
                CustomBarcodeGeneratorUtils::SetPosCodeStyle(gen, parameters->get_PosCodeStyle(), parameters->get_BarcodeValue());
            }

            return CustomBarcodeGeneratorUtils::ConvertImageToWord(gen->GenerateBarCodeImage());
        }
        catch (System::Exception& e)
        {
            return CustomBarcodeGeneratorUtils::ConvertImageToWord(CustomBarcodeGeneratorUtils::DrawErrorImage(e->get_Message()));
        }
    }

    SharedPtr<System::IO::Stream> GetOldBarcodeImage(SharedPtr<Aspose::Words::Fields::BarcodeParameters> parameters) override
    {
        throw System::NotImplementedException();
    }
};

} // namespace DocsExamples

#endif // ASPOSE_BARCODE_AVAILABLE
