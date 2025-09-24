// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "TestUtil.h"

#include <testing/test_predicates.h>
#include <system/text/string_builder.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/platform_id.h>
#include <system/operating_system.h>
#include <system/linq/enumerable.h>
#include <system/io/stream_reader.h>
#include <system/io/path.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file.h>
#include <system/io/binary_reader.h>
#include <system/exceptions.h>
#include <system/environment.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <gtest/gtest.h>
#include <drawing/rectangle.h>
#include <drawing/imaging/metafile_header.h>
#include <drawing/imaging/metafile.h>
#include <drawing/image.h>
#include <drawing/color.h>
#include <drawing/bitmap.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Notes;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1545707249u, ::Aspose::Words::ApiExamples::TestUtil, ThisTypeBaseTypesInfo);

void TestUtil::VerifyImage(int32_t expectedWidth, int32_t expectedHeight, System::String filename)
{
    System::PlatformID pid = System::Environment::get_OSVersion().get_Platform();
    bool isWindows = (pid == System::PlatformID::Win32NT) || (pid == System::PlatformID::Win32S) || (pid == System::PlatformID::Win32Windows) || (pid == System::PlatformID::WinCE);
    System::String ext = System::IO::Path::GetExtension(filename).ToLower();
    
    if (isWindows && (ext == u".emf" || ext == u".wmf"))
    {
        {
            auto metafile = System::MakeObject<System::Drawing::Imaging::Metafile>(filename);
            System::Drawing::Rectangle bounds = metafile->GetMetafileHeader()->get_Bounds();
            ASSERT_NEAR(expectedWidth, bounds.get_Width(), 1);
            ASSERT_NEAR(expectedHeight, bounds.get_Height(), 1);
        }
    }
    else if (ext == u".emf")
    {
        System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> emfDimensions = GetEmfDimensions(filename);
        ASSERT_NEAR(expectedWidth, emfDimensions->idx_get(u"width"), 1);
        ASSERT_NEAR(expectedHeight, emfDimensions->idx_get(u"height"), 1);
    }
    else if (ext == u".wmf")
    {
        System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> wmfDimensions = GetWmfDimensions(filename);
        ASSERT_NEAR(expectedWidth, wmfDimensions->idx_get(u"width"), 1);
        ASSERT_NEAR(expectedHeight, wmfDimensions->idx_get(u"height"), 1);
    }
    else
    {
        {
            auto image = System::Drawing::Image::FromFile(filename);
            ASSERT_NEAR(expectedWidth, image->get_Width(), 1);
            ASSERT_NEAR(expectedHeight, image->get_Height(), 1);
        }
    }
}

System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> TestUtil::GetEmfDimensions(System::String filePath)
{
    {
        auto stream = System::IO::File::OpenRead(filePath);
        {
            auto reader = System::MakeObject<System::IO::BinaryReader>(stream);
            // Skip EMF header (first 8 bytes).
            stream->set_Position(8);
            
            // Read bounding rectangle (4 x Int32: left, top, right, bottom).
            int32_t left = reader->ReadInt32();
            int32_t top = reader->ReadInt32();
            int32_t right = reader->ReadInt32();
            int32_t bottom = reader->ReadInt32();
            
            System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> emfDimensions = System::MakeObject<System::Collections::Generic::Dictionary<System::String, int32_t>>();
            emfDimensions->Add(u"width", right - left);
            emfDimensions->Add(u"height", bottom - top);
            
            return emfDimensions;
        }
    }
}

System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> TestUtil::GetWmfDimensions(System::String filePath)
{
    {
        auto stream = System::IO::File::OpenRead(filePath);
        {
            auto reader = System::MakeObject<System::IO::BinaryReader>(stream);
            // WMF header (16 bytes).
            // Skip first 10 bytes (header + version).
            stream->set_Position(10);
            
            // Read dimensions in 16-bit integers (0.01mm units).
            int16_t left = reader->ReadInt16();
            int16_t top = reader->ReadInt16();
            int16_t right = reader->ReadInt16();
            int16_t bottom = reader->ReadInt16();
            
            // Convert to pixels (96 DPI approximation).
            const double unitsPerInch = 2540.0;
            // WMF uses 0.01mm units.
            int32_t width = (int32_t)((right - left) / unitsPerInch * 96);
            int32_t height = (int32_t)((bottom - top) / unitsPerInch * 96);
            
            System::SharedPtr<System::Collections::Generic::Dictionary<System::String, int32_t>> wmfDimensions = System::MakeObject<System::Collections::Generic::Dictionary<System::String, int32_t>>();
            wmfDimensions->Add(u"width", width);
            wmfDimensions->Add(u"height", height);
            
            return wmfDimensions;
        }
    }
}

void TestUtil::ImageContainsTransparency(System::String filename)
{
    {
        auto image = System::Drawing::Image::FromFile(filename);
        {
            auto bitmap = System::MakeObject<System::Drawing::Bitmap>(image);
            for (int32_t x = 0; x < bitmap->get_Width(); x++)
            {
                for (int32_t y = 0; y < bitmap->get_Height(); y++)
                {
                    System::Drawing::Color pixel = bitmap->GetPixel(x, y);
                    if (pixel.get_A() != 255)
                    {
                        return;
                    }
                    // Transparency found.
                }
            }
        }
    }
    FAIL() << (System::String::Format(u"The image from \"{0}\" does not contain any transparency.", filename));
}

void TestUtil::MailMergeMatchesArray(System::ArrayPtr<System::ArrayPtr<System::String>> expectedResult, System::SharedPtr<Aspose::Words::Document> doc, bool onePagePerRow)
{
    try
    {
        if (onePagePerRow)
        {
            System::ArrayPtr<System::String> docTextByPages = doc->GetText().Trim().Split(System::MakeArray<System::String>({Aspose::Words::ControlChar::PageBreak()}), System::StringSplitOptions::RemoveEmptyEntries);
            
            for (int32_t i = 0; i < expectedResult->get_Length(); i++)
            {
                for (int32_t j = 0; j < expectedResult[0]->get_Length(); j++)
                {
                    if (!docTextByPages[i].Contains(expectedResult[i][j]))
                    {
                        throw System::ArgumentException(expectedResult[i][j]);
                    }
                }
            }
        }
        else
        {
            System::String docText = doc->GetText();
            
            for (int32_t i = 0; i < expectedResult->get_Length(); i++)
            {
                for (int32_t j = 0; j < expectedResult[0]->get_Length(); j++)
                {
                    if (!docText.Contains(expectedResult[i][j]))
                    {
                        throw System::ArgumentException(expectedResult[i][j]);
                    }
                }
            }
            
        }
    }
    catch (System::ArgumentException& e)
    {
        FAIL() << (System::String::Format(u"String \"{0}\" not found in {1}.", e->get_Message(), (doc->get_OriginalFileName() == nullptr ? u"a virtual document" : doc->get_OriginalFileName().Split(u'\\')->LINQ_Last())));
    }
    
}

void TestUtil::FileContainsString(System::String expected, System::String filename)
{
    {
        System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::FileStream>(filename, System::IO::FileMode::Open);
        StreamContainsString(expected, stream);
    }
}

void TestUtil::StreamContainsString(System::String expected, System::SharedPtr<System::IO::Stream> stream)
{
    System::ArrayPtr<char16_t> expectedSequence = expected.ToCharArray();
    
    int64_t sequenceMatchLength = 0;
    while (stream->get_Position() < stream->get_Length())
    {
        if ((char16_t)stream->ReadByte() == expectedSequence[sequenceMatchLength])
        {
            sequenceMatchLength++;
        }
        else
        {
            sequenceMatchLength = 0;
        }
        
        if (sequenceMatchLength >= expectedSequence->get_Length())
        {
            return;
        }
    }
    
    FAIL() << (System::String::Format(u"String \"{0}\" not found in the provided source.", (expected.get_Length() <= 100 ? expected : expected.Substring(0, 100) + u"...")));
}

void TestUtil::VerifyField(Aspose::Words::Fields::FieldType expectedType, System::String expectedFieldCode, System::String expectedResult, System::SharedPtr<Aspose::Words::Fields::Field> field)
{
    ASSERT_EQ(expectedType, field->get_Type());
    ASSERT_EQ(expectedFieldCode, field->GetFieldCode(true));
    ASSERT_EQ(expectedResult, field->get_Result());
}

void TestUtil::VerifyField(Aspose::Words::Fields::FieldType expectedType, System::String expectedFieldCode, System::DateTime expectedResult, System::SharedPtr<Aspose::Words::Fields::Field> field, System::TimeSpan delta)
{
    ASSERT_EQ(expectedType, field->get_Type());
    ASSERT_EQ(expectedFieldCode, field->GetFieldCode(true));
    System::DateTime actual = System::DateTime::get_Now();
    try
    {
        actual = System::DateTime::Parse(field->get_Result());
    }
    catch (System::Exception& )
    {
        FAIL();
    }
    
    if (field->get_Type() == Aspose::Words::Fields::FieldType::FieldTime)
    {
        VerifyDate(expectedResult, actual, delta);
    }
    else
    {
        VerifyDate(expectedResult.get_Date(), actual, delta);
    }
}

void TestUtil::VerifyDate(System::DateTime expected, System::DateTime actual, System::TimeSpan delta)
{
    ASSERT_TRUE(expected - actual <= delta);
}

void TestUtil::FieldsAreNested(System::SharedPtr<Aspose::Words::Fields::Field> innerField, System::SharedPtr<Aspose::Words::Fields::Field> outerField)
{
    System::SharedPtr<Aspose::Words::CompositeNode> innerFieldParent = innerField->get_Start()->get_ParentNode();
    
    ASSERT_TRUE(innerFieldParent == outerField->get_Start()->get_ParentNode());
    ASSERT_TRUE(innerFieldParent->GetChildNodes(Aspose::Words::NodeType::Any, false)->IndexOf(innerField->get_Start()) > innerFieldParent->GetChildNodes(Aspose::Words::NodeType::Any, false)->IndexOf(outerField->get_Start()));
    ASSERT_TRUE(innerFieldParent->GetChildNodes(Aspose::Words::NodeType::Any, false)->IndexOf(innerField->get_End()) < innerFieldParent->GetChildNodes(Aspose::Words::NodeType::Any, false)->IndexOf(outerField->get_End()));
}

void TestUtil::VerifyImageInShape(int32_t expectedWidth, int32_t expectedHeight, Aspose::Words::Drawing::ImageType expectedImageType, System::SharedPtr<Aspose::Words::Drawing::Shape> imageShape)
{
    ASSERT_TRUE(imageShape->get_HasImage());
    ASSERT_EQ(expectedImageType, imageShape->get_ImageData()->get_ImageType());
    ASSERT_EQ(expectedWidth, imageShape->get_ImageData()->get_ImageSize()->get_WidthPixels());
    ASSERT_EQ(expectedHeight, imageShape->get_ImageData()->get_ImageSize()->get_HeightPixels());
}

void TestUtil::VerifyFootnote(Aspose::Words::Notes::FootnoteType expectedFootnoteType, bool expectedIsAuto, System::String expectedReferenceMark, System::String expectedContents, System::SharedPtr<Aspose::Words::Notes::Footnote> footnote)
{
    ASSERT_EQ(expectedFootnoteType, footnote->get_FootnoteType());
    ASPOSE_ASSERT_EQ(expectedIsAuto, footnote->get_IsAuto());
    ASSERT_EQ(expectedReferenceMark, footnote->get_ReferenceMark());
    ASSERT_EQ(expectedContents, footnote->ToString(Aspose::Words::SaveFormat::Text).Trim());
}

void TestUtil::VerifyListLevel(System::String expectedListFormat, double expectedNumberPosition, Aspose::Words::NumberStyle expectedNumberStyle, System::SharedPtr<Aspose::Words::Lists::ListLevel> listLevel)
{
    ASSERT_EQ(expectedListFormat, listLevel->get_NumberFormat());
    ASPOSE_ASSERT_EQ(expectedNumberPosition, listLevel->get_NumberPosition());
    ASSERT_EQ(expectedNumberStyle, listLevel->get_NumberStyle());
}

void TestUtil::CopyStream(System::SharedPtr<System::IO::Stream> srcStream, System::SharedPtr<System::IO::Stream> dstStream)
{
    if (srcStream == nullptr)
    {
        throw System::ArgumentNullException(u"srcStream");
    }
    if (dstStream == nullptr)
    {
        throw System::ArgumentNullException(u"dstStream");
    }
    
    auto buf = System::MakeArray<uint8_t>(65536, 0);
    while (true)
    {
        int32_t bytesRead = srcStream->Read(buf, 0, buf->get_Length());
        // Read returns 0 when reached end of stream
        // Checking for negative too to make it conceptually close to Java
        if (bytesRead <= 0)
        {
            break;
        }
        dstStream->Write(buf, 0, bytesRead);
    }
}

System::String TestUtil::DumpArray(System::ArrayPtr<uint8_t> data, int32_t start, int32_t count)
{
    if (data == nullptr)
    {
        return u"Null";
    }
    
    System::SharedPtr<System::Text::StringBuilder> builder = System::MakeObject<System::Text::StringBuilder>();
    while (count > 0)
    {
        builder->AppendFormat(u"{0:X2} ", data[start]);
        start++;
        count--;
    }
    return builder->ToString();
}

void TestUtil::VerifyTabStop(double expectedPosition, Aspose::Words::TabAlignment expectedTabAlignment, Aspose::Words::TabLeader expectedTabLeader, bool isClear, System::SharedPtr<Aspose::Words::TabStop> tabStop)
{
    ASPOSE_ASSERT_EQ(expectedPosition, tabStop->get_Position());
    ASSERT_EQ(expectedTabAlignment, tabStop->get_Alignment());
    ASSERT_EQ(expectedTabLeader, tabStop->get_Leader());
    ASPOSE_ASSERT_EQ(isClear, tabStop->get_IsClear());
}

void TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType expectedShapeType, System::String expectedName, double expectedWidth, double expectedHeight, double expectedTop, double expectedLeft, System::SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
    ASSERT_EQ(expectedShapeType, shape->get_ShapeType());
    ASSERT_EQ(expectedName, shape->get_Name());
    ASPOSE_ASSERT_EQ(expectedWidth, shape->get_Width());
    ASPOSE_ASSERT_EQ(expectedHeight, shape->get_Height());
    ASPOSE_ASSERT_EQ(expectedTop, shape->get_Top());
    ASPOSE_ASSERT_EQ(expectedLeft, shape->get_Left());
}

void TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow expectedLayoutFlow, bool expectedFitShapeToText, Aspose::Words::Drawing::TextBoxWrapMode expectedTextBoxWrapMode, double marginTop, double marginBottom, double marginLeft, double marginRight, System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox)
{
    ASSERT_EQ(expectedLayoutFlow, textBox->get_LayoutFlow());
    ASPOSE_ASSERT_EQ(expectedFitShapeToText, textBox->get_FitShapeToText());
    ASSERT_EQ(expectedTextBoxWrapMode, textBox->get_TextBoxWrapMode());
    ASPOSE_ASSERT_EQ(marginTop, textBox->get_InternalMarginTop());
    ASPOSE_ASSERT_EQ(marginBottom, textBox->get_InternalMarginBottom());
    ASPOSE_ASSERT_EQ(marginLeft, textBox->get_InternalMarginLeft());
    ASPOSE_ASSERT_EQ(marginRight, textBox->get_InternalMarginRight());
}

void TestUtil::VerifyEditableRange(int32_t expectedId, System::String expectedEditorUser, Aspose::Words::EditorType expectedEditorGroup, System::SharedPtr<Aspose::Words::EditableRange> editableRange)
{
    ASSERT_EQ(expectedId, editableRange->get_Id());
    ASSERT_EQ(expectedEditorUser, editableRange->get_SingleUser());
    ASSERT_EQ(expectedEditorGroup, editableRange->get_EditorGroup());
}

System::SharedPtr<System::Text::Encoding> TestUtil::GetEncoding(System::String filename)
{
    {
        auto streamReader = System::MakeObject<System::IO::StreamReader>(filename, true);
        streamReader->Read();
        return streamReader->get_CurrentEncoding();
    }
}

System::ArrayPtr<uint8_t> TestUtil::ImageToByteArray(System::String imagePath)
{
    return System::IO::File::ReadAllBytes(imagePath);
}

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
