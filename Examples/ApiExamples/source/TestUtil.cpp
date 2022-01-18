// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "TestUtil.h"

#include <cstdint>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageSize.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <drawing/bitmap.h>
#include <drawing/color.h>
#include <drawing/image.h>
#include <net/http_web_request.h>
#include <net/http_web_response.h>
#include <net/web_request.h>
#include <net/web_response.h>
#include <system/details/dispose_guard.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/linq/enumerable.h>
#include <system/string.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/string_builder.h>
#include <gtest/gtest.h>
#include <testing/test_predicates.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Notes;
namespace ApiExamples {

void TestUtil::VerifyImage(int32_t expectedWidth, int32_t expectedHeight, System::String filename)
{
    {
        auto fileStream = System::MakeObject<System::IO::FileStream>(filename, System::IO::FileMode::Open);
        VerifyImage(expectedWidth, expectedHeight, fileStream);
    }
}

void TestUtil::VerifyImage(int32_t expectedWidth, int32_t expectedHeight, System::SharedPtr<System::IO::Stream> imageStream)
{
    {
        System::SharedPtr<System::Drawing::Image> image = System::Drawing::Image::FromStream(imageStream);
        ASSERT_EQ(expectedWidth, image->get_Width());
        ASSERT_EQ(expectedHeight, image->get_Height());
    }
}

void TestUtil::ImageContainsTransparency(System::String filename)
{
    {
        auto bitmap = System::DynamicCast<System::Drawing::Bitmap>(System::Drawing::Image::FromFile(filename));
        for (int32_t x = 0; x < bitmap->get_Width(); x++)
        {
            for (int32_t y = 0; y < bitmap->get_Height(); y++)
            {
                if (bitmap->GetPixel(x, y).get_A() != 255)
                {
                    return;
                }
            }
        }
    }

    FAIL() << (System::String::Format(u"The image from \"{0}\" does not contain any transparency.", filename));
}

void TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode expectedHttpStatusCode, System::String webAddress)
{
    auto request = System::DynamicCast<System::Net::HttpWebRequest>(System::Net::WebRequest::Create(webAddress));
    request->set_Method(u"HEAD");

    ASSERT_EQ(expectedHttpStatusCode, (System::DynamicCast<System::Net::HttpWebResponse>(request->GetResponse()))->get_StatusCode());
}

void TestUtil::TableMatchesQueryResult(System::SharedPtr<Tables::Table> expectedResult, System::String dbFilename, System::String sqlQuery)
{
}

void TestUtil::MailMergeMatchesQueryResultMultiple(System::String dbFilename, System::ArrayPtr<System::String> sqlQueries, System::SharedPtr<Document> doc,
                                                   bool onePagePerRow)
{
    for (System::String query : sqlQueries)
    {
        MailMergeMatchesQueryResult(dbFilename, query, doc, onePagePerRow);
    }
}

void TestUtil::MailMergeMatchesQueryResult(System::String dbFilename, System::String sqlQuery, System::SharedPtr<Document> doc, bool onePagePerRow)
{
}

void TestUtil::MailMergeMatchesArray(System::ArrayPtr<System::ArrayPtr<System::String>> expectedResult, System::SharedPtr<Document> doc, bool onePagePerRow)
{
    try
    {
        if (onePagePerRow)
        {
            System::ArrayPtr<System::String> docTextByPages =
                doc->GetText().Trim().Split(System::MakeArray<System::String>({ControlChar::PageBreak()}), System::StringSplitOptions::RemoveEmptyEntries);

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
        FAIL() << (System::String::Format(u"String \"{0}\" not found in {1}.", e->get_Message(),
                                          (doc->get_OriginalFileName() == nullptr
                                               ? u"a virtual document"
                                               : doc->get_OriginalFileName().Split(System::MakeArray<char16_t>({u'\\'}))->LINQ_Last())));
    }
}

void TestUtil::FileContainsString(System::String expected, System::String filename)
{
    if (!IsRunningOnMono())
    {
        {
            System::SharedPtr<System::IO::Stream> stream = System::MakeObject<System::IO::FileStream>(filename, System::IO::FileMode::Open);
            StreamContainsString(expected, stream);
        }
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

    FAIL() << (System::String::Format(u"String \"{0}\" not found in the provided source.",
                                      (expected.get_Length() <= 100 ? expected : expected.Substring(0, 100) + u"...")));
}

void TestUtil::VerifyField(FieldType expectedType, System::String expectedFieldCode, System::String expectedResult, System::SharedPtr<Field> field)
{
    ASSERT_EQ(expectedType, field->get_Type());
    ASSERT_EQ(expectedFieldCode, field->GetFieldCode(true));
    ASSERT_EQ(expectedResult, field->get_Result());
}

void TestUtil::VerifyField(FieldType expectedType, System::String expectedFieldCode, System::DateTime expectedResult, System::SharedPtr<Field> field,
                           System::TimeSpan delta)
{
}

void TestUtil::VerifyDate(System::DateTime expected, System::DateTime actual, System::TimeSpan delta)
{
    ASSERT_TRUE(expected - actual <= delta);
}

void TestUtil::FieldsAreNested(System::SharedPtr<Field> innerField, System::SharedPtr<Field> outerField)
{
    System::SharedPtr<CompositeNode> innerFieldParent = innerField->get_Start()->get_ParentNode();

    ASSERT_TRUE(innerFieldParent == outerField->get_Start()->get_ParentNode());
    ASSERT_TRUE(innerFieldParent->get_ChildNodes()->IndexOf(innerField->get_Start()) > innerFieldParent->get_ChildNodes()->IndexOf(outerField->get_Start()));
    ASSERT_TRUE(innerFieldParent->get_ChildNodes()->IndexOf(innerField->get_End()) < innerFieldParent->get_ChildNodes()->IndexOf(outerField->get_End()));
}

void TestUtil::VerifyImageInShape(int32_t expectedWidth, int32_t expectedHeight, ImageType expectedImageType, System::SharedPtr<Shape> imageShape)
{
    ASSERT_TRUE(imageShape->get_HasImage());
    ASSERT_EQ(expectedImageType, imageShape->get_ImageData()->get_ImageType());
    ASSERT_EQ(expectedWidth, imageShape->get_ImageData()->get_ImageSize()->get_WidthPixels());
    ASSERT_EQ(expectedHeight, imageShape->get_ImageData()->get_ImageSize()->get_HeightPixels());
}

void TestUtil::VerifyFootnote(FootnoteType expectedFootnoteType, bool expectedIsAuto, System::String expectedReferenceMark, System::String expectedContents,
                              System::SharedPtr<Footnote> footnote)
{
    ASSERT_EQ(expectedFootnoteType, footnote->get_FootnoteType());
    ASPOSE_ASSERT_EQ(expectedIsAuto, footnote->get_IsAuto());
    ASSERT_EQ(expectedReferenceMark, footnote->get_ReferenceMark());
    ASSERT_EQ(expectedContents, footnote->ToString(SaveFormat::Text).Trim());
}

void TestUtil::VerifyListLevel(System::String expectedListFormat, double expectedNumberPosition, NumberStyle expectedNumberStyle,
                               System::SharedPtr<ListLevel> listLevel)
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

    auto builder = System::MakeObject<System::Text::StringBuilder>();
    while (count > 0)
    {
        builder->AppendFormat(u"{0:X2} ", data[start]);
        start++;
        count--;
    }
    return builder->ToString();
}

void TestUtil::VerifyTabStop(double expectedPosition, TabAlignment expectedTabAlignment, TabLeader expectedTabLeader, bool isClear,
                             System::SharedPtr<TabStop> tabStop)
{
    ASPOSE_ASSERT_EQ(expectedPosition, tabStop->get_Position());
    ASSERT_EQ(expectedTabAlignment, tabStop->get_Alignment());
    ASSERT_EQ(expectedTabLeader, tabStop->get_Leader());
    ASPOSE_ASSERT_EQ(isClear, tabStop->get_IsClear());
}

void TestUtil::VerifyShape(ShapeType expectedShapeType, System::String expectedName, double expectedWidth, double expectedHeight, double expectedTop,
                           double expectedLeft, System::SharedPtr<Shape> shape)
{
    ASSERT_EQ(expectedShapeType, shape->get_ShapeType());
    ASSERT_EQ(expectedName, shape->get_Name());
    ASPOSE_ASSERT_EQ(expectedWidth, shape->get_Width());
    ASPOSE_ASSERT_EQ(expectedHeight, shape->get_Height());
    ASPOSE_ASSERT_EQ(expectedTop, shape->get_Top());
    ASPOSE_ASSERT_EQ(expectedLeft, shape->get_Left());
}

void TestUtil::VerifyTextBox(LayoutFlow expectedLayoutFlow, bool expectedFitShapeToText, TextBoxWrapMode expectedTextBoxWrapMode, double marginTop,
                             double marginBottom, double marginLeft, double marginRight, System::SharedPtr<TextBox> textBox)
{
    ASSERT_EQ(expectedLayoutFlow, textBox->get_LayoutFlow());
    ASPOSE_ASSERT_EQ(expectedFitShapeToText, textBox->get_FitShapeToText());
    ASSERT_EQ(expectedTextBoxWrapMode, textBox->get_TextBoxWrapMode());
    ASPOSE_ASSERT_EQ(marginTop, textBox->get_InternalMarginTop());
    ASPOSE_ASSERT_EQ(marginBottom, textBox->get_InternalMarginBottom());
    ASPOSE_ASSERT_EQ(marginLeft, textBox->get_InternalMarginLeft());
    ASPOSE_ASSERT_EQ(marginRight, textBox->get_InternalMarginRight());
}

void TestUtil::VerifyEditableRange(int32_t expectedId, System::String expectedEditorUser, EditorType expectedEditorGroup,
                                   System::SharedPtr<EditableRange> editableRange)
{
    ASSERT_EQ(expectedId, editableRange->get_Id());
    ASSERT_EQ(expectedEditorUser, editableRange->get_SingleUser());
    ASSERT_EQ(expectedEditorGroup, editableRange->get_EditorGroup());
}

} // namespace ApiExamples
