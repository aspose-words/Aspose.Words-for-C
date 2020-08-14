// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "TestUtil.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/exceptions.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/array.h>
#include <net/web_response.h>
#include <net/web_request.h>
#include <net/http_web_response.h>
#include <net/http_web_request.h>
#include <net/http_status_code.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/TabStop.h>
#include <Aspose.Words.Cpp/Model/Text/TabLeader.h>
#include <Aspose.Words.Cpp/Model/Text/TabAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBoxWrapMode.h>
#include <Aspose.Words.Cpp/Model/Drawing/TextBox.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/LayoutFlow.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Lists;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2267053751u, ::ApiExamples::TestUtil, ThisTypeBaseTypesInfo);

int32_t TestUtil::get_FileInfoLengthDelta()
{
    return 200;
}

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
    }
}

void TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode expectedHttpStatusCode, System::String webAddress)
{
    auto request = System::DynamicCast<System::Net::HttpWebRequest>(System::Net::WebRequest::Create(webAddress));
    request->set_Method(u"HEAD");

    ASSERT_EQ(expectedHttpStatusCode, (System::DynamicCast<System::Net::HttpWebResponse>(request->GetResponse()))->get_StatusCode());
}

void TestUtil::TableMatchesQueryResult(System::SharedPtr<Aspose::Words::Tables::Table> expectedResult, System::String dbFilename, System::String sqlQuery)
{
}

void TestUtil::MailMergeMatchesQueryResultMultiple(System::String dbFilename, System::ArrayPtr<System::String> sqlQueries, System::SharedPtr<Aspose::Words::Document> doc, bool onePagePerRow)
{
    for (System::String query : sqlQueries)
    {
        MailMergeMatchesQueryResult(dbFilename, query, doc, onePagePerRow);
    }

}

void TestUtil::MailMergeMatchesQueryResult(System::String dbFilename, System::String sqlQuery, System::SharedPtr<Aspose::Words::Document> doc, bool onePagePerRow)
{
}

void TestUtil::MailMergeMatchesArray(System::ArrayPtr<System::ArrayPtr<System::String>> expectedResult, System::SharedPtr<Aspose::Words::Document> doc, bool onePagePerRow)
{
    try
    {
        if (onePagePerRow)
        {
            System::ArrayPtr<System::String> docTextByPages = doc->GetText().Trim().Split(System::MakeArray<System::String>({ControlChar::PageBreak()}), System::StringSplitOptions::RemoveEmptyEntries);

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
        FAIL() << (System::String::Format(u"String \"{0}\" not found in {1}.",e->get_Message(),(doc->get_OriginalFileName() == nullptr ? u"a virtual document" : doc->get_OriginalFileName().Split(System::MakeArray<char16_t>({u'\\'}))->LINQ_Last()))).ToUtf8String();
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

    FAIL() << (System::String::Format(u"String \"{0}\" not found in the provided source.",(expected.get_Length() <= 100 ? expected : expected.Substring(0, 100) + u"..."))).ToUtf8String();
}

void TestUtil::VerifyField(Aspose::Words::Fields::FieldType expectedType, System::String expectedFieldCode, System::String expectedResult, System::SharedPtr<Aspose::Words::Fields::Field> field)
{
}

void TestUtil::VerifyField(Aspose::Words::Fields::FieldType expectedType, System::String expectedFieldCode, System::DateTime expectedResult, System::SharedPtr<Aspose::Words::Fields::Field> field, System::TimeSpan delta)
{
}

void TestUtil::VerifyDate(System::DateTime expected, System::DateTime actual, System::TimeSpan delta)
{
    ASSERT_TRUE(expected - actual <= delta);
}

void TestUtil::FieldsAreNested(System::SharedPtr<Aspose::Words::Fields::Field> innerField, System::SharedPtr<Aspose::Words::Fields::Field> outerField)
{
    System::SharedPtr<CompositeNode> innerFieldParent = innerField->get_Start()->get_ParentNode();

    ASSERT_TRUE(innerFieldParent == outerField->get_Start()->get_ParentNode());
    ASSERT_TRUE(innerFieldParent->get_ChildNodes()->IndexOf(innerField->get_Start()) > innerFieldParent->get_ChildNodes()->IndexOf(outerField->get_Start()));
    ASSERT_TRUE(innerFieldParent->get_ChildNodes()->IndexOf(innerField->get_End()) < innerFieldParent->get_ChildNodes()->IndexOf(outerField->get_End()));
}

void TestUtil::VerifyImageInShape(int32_t expectedWidth, int32_t expectedHeight, Aspose::Words::Drawing::ImageType expectedImageType, System::SharedPtr<Aspose::Words::Drawing::Shape> imageShape)
{
}

void TestUtil::VerifyFootnote(Aspose::Words::FootnoteType expectedFootnoteType, bool expectedIsAuto, System::String expectedReferenceMark, System::String expectedContents, System::SharedPtr<Aspose::Words::Footnote> footnote)
{
}

void TestUtil::VerifyListLevel(System::String expectedListFormat, double expectedNumberPosition, Aspose::Words::NumberStyle expectedNumberStyle, System::SharedPtr<Aspose::Words::Lists::ListLevel> listLevel)
{
}

void TestUtil::VerifyTabStop(double expectedPosition, Aspose::Words::TabAlignment expectedTabAlignment, Aspose::Words::TabLeader expectedTabLeader, bool isClear, System::SharedPtr<Aspose::Words::TabStop> tabStop)
{
}

void TestUtil::VerifyShape(Aspose::Words::Drawing::ShapeType expectedShapeType, System::String expectedName, double expectedWidth, double expectedHeight, double expectedTop, double expectedLeft, System::SharedPtr<Aspose::Words::Drawing::Shape> shape)
{
}

void TestUtil::VerifyTextBox(Aspose::Words::Drawing::LayoutFlow expectedLayoutFlow, bool expectedFitShapeToText, Aspose::Words::Drawing::TextBoxWrapMode expectedTextBoxWrapMode, double marginTop, double marginBottom, double marginLeft, double marginRight, System::SharedPtr<Aspose::Words::Drawing::TextBox> textBox)
{
}

} // namespace ApiExamples
