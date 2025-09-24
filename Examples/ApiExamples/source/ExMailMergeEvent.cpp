// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMailMergeEvent.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/primitive_types.h>
#include <system/object_ext.h>
#include <system/io/stream.h>
#include <system/io/memory_stream.h>
#include <system/char.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>

#include "TestUtil.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::MailMerging;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(373933079u, ::Aspose::Words::ApiExamples::ExMailMergeEvent::HandleMergeFieldInsertHtml, ThisTypeBaseTypesInfo);

void ExMailMergeEvent::HandleMergeFieldInsertHtml::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
{
    if (args->get_DocumentFieldName().StartsWith(u"html_") && args->get_Field()->GetFieldCode().Contains(u"\\b"))
    {
        // Add parsed HTML data to the document's body.
        auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(args->get_Document());
        builder->MoveToMergeField(args->get_DocumentFieldName());
        builder->InsertHtml(System::ExplicitCast<System::String>(args->get_FieldValue()));
        
        // Since we have already inserted the merged content manually,
        // we will not need to respond to this event by returning content via the "Text" property. 
        args->set_Text(System::String::Empty);
    }
}

void ExMailMergeEvent::HandleMergeFieldInsertHtml::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    // Do nothing.
}

RTTI_INFO_IMPL_HASH(3670203834u, ::Aspose::Words::ApiExamples::ExMailMergeEvent::FieldValueMergingCallback, ThisTypeBaseTypesInfo);

void ExMailMergeEvent::FieldValueMergingCallback::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> e)
{
    if (e->get_FieldName().StartsWith(u"text_"))
    {
        e->set_FieldValue(System::ExplicitCast<System::Object>(System::String::Format(u"Merge value for \"{0}\": {1}", e->get_FieldName(), System::ExplicitCast<System::String>(e->get_FieldValue()))));
    }
    else if (e->get_FieldName().StartsWith(u"numeric_"))
    {
        e->set_FieldValue(System::ExplicitCast<System::Object>(System::ExplicitCast<int32_t>(e->get_FieldValue()) * 1000));
    }
}

void ExMailMergeEvent::FieldValueMergingCallback::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> e)
{
    // Do nothing.
}

RTTI_INFO_IMPL_HASH(2239537199u, ::Aspose::Words::ApiExamples::ExMailMergeEvent::HandleMergeFieldInsertCheckBox, ThisTypeBaseTypesInfo);

void ExMailMergeEvent::HandleMergeFieldInsertCheckBox::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
{
    if (args->get_DocumentFieldName() == u"CourseName")
    {
        ASSERT_EQ(u"StudentCourse", args->get_TableName());
        
        auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(args->get_Document());
        builder->MoveToMergeField(args->get_FieldName());
        builder->InsertCheckBox(args->get_DocumentFieldName() + mCheckBoxCount, false, 0);
        
        System::String fieldValue = System::ObjectExt::ToString(args->get_FieldValue());
        
        // In this case, for every record index 'n', the corresponding field value is "Course n".
        ASPOSE_ASSERT_EQ(System::Char::GetNumericValue(fieldValue[7]), args->get_RecordIndex());
        
        builder->Write(fieldValue);
        mCheckBoxCount++;
    }
}

void ExMailMergeEvent::HandleMergeFieldInsertCheckBox::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    // Do nothing.
}

ExMailMergeEvent::HandleMergeFieldInsertCheckBox::HandleMergeFieldInsertCheckBox() : mCheckBoxCount(0)
{
}

RTTI_INFO_IMPL_HASH(977494397u, ::Aspose::Words::ApiExamples::ExMailMergeEvent::HandleMergeFieldAlternatingRows, ThisTypeBaseTypesInfo);

void ExMailMergeEvent::HandleMergeFieldAlternatingRows::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
{
    if (mBuilder == nullptr)
    {
        mBuilder = System::MakeObject<Aspose::Words::DocumentBuilder>(args->get_Document());
    }
    
    // This is true of we are on the first column, which means we have moved to a new row.
    if (args->get_FieldName() == u"CompanyName")
    {
        System::Drawing::Color rowColor = IsOdd(mRowIdx) ? System::Drawing::Color::FromArgb(213, 227, 235) : System::Drawing::Color::FromArgb(242, 242, 242);
        
        for (int32_t colIdx = 0; colIdx < 4; colIdx++)
        {
            mBuilder->MoveToCell(0, mRowIdx, colIdx, 0);
            mBuilder->get_CellFormat()->get_Shading()->set_BackgroundPatternColor(rowColor);
        }
        
        mRowIdx++;
    }
}

void ExMailMergeEvent::HandleMergeFieldAlternatingRows::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    // Do nothing.
}

ExMailMergeEvent::HandleMergeFieldAlternatingRows::HandleMergeFieldAlternatingRows() : mRowIdx(0)
{
}

RTTI_INFO_IMPL_HASH(809343725u, ::Aspose::Words::ApiExamples::ExMailMergeEvent::HandleMergeImageFieldFromBlob, ThisTypeBaseTypesInfo);

void ExMailMergeEvent::HandleMergeImageFieldFromBlob::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
{
    // Do nothing.
}

void ExMailMergeEvent::HandleMergeImageFieldFromBlob::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> e)
{
    auto imageStream = System::MakeObject<System::IO::MemoryStream>(System::ExplicitCast<System::Array<uint8_t>>(e->get_FieldValue()));
    e->set_ImageStream(imageStream);
}


RTTI_INFO_IMPL_HASH(529416553u, ::Aspose::Words::ApiExamples::ExMailMergeEvent, ThisTypeBaseTypesInfo);

bool ExMailMergeEvent::IsOdd(int32_t value)
{
    return System::Equals((value / 2 * 2), value);
}


namespace gtest_test
{

class ExMailMergeEvent : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExMailMergeEvent> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExMailMergeEvent>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExMailMergeEvent> ExMailMergeEvent::s_instance;

} // namespace gtest_test

void ExMailMergeEvent::MergeHtml()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertField(u"MERGEFIELD  html_Title  \\b Content");
    builder->InsertField(u"MERGEFIELD  html_Body  \\b Content");
    
    System::ArrayPtr<System::SharedPtr<System::Object>> mergeData = System::MakeArray<System::SharedPtr<System::Object>>({System::ExplicitCast<System::Object>(System::String(u"<html>") + u"<h1>" + u"<span style=\"color: #0000ff; font-family: Arial;\">Hello World!</span>" + u"</h1>" + u"</html>"), System::ExplicitCast<System::Object>(System::String(u"<html>") + u"<blockquote>" + u"<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>" + u"</blockquote>" + u"</html>")});
    
    doc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExMailMergeEvent::HandleMergeFieldInsertHtml>());
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"html_Title", u"html_Body"}), mergeData);
    
    doc->Save(get_ArtifactsDir() + u"MailMergeEvent.MergeHtml.docx");
}

namespace gtest_test
{

TEST_F(ExMailMergeEvent, MergeHtml)
{
    s_instance->MergeHtml();
}

} // namespace gtest_test

void ExMailMergeEvent::FieldFormats()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert some MERGEFIELDs with format switches that will edit the values they will receive during a mail merge.
    builder->InsertField(u"MERGEFIELD text_Field1 \\* Caps", nullptr);
    builder->Write(u", ");
    builder->InsertField(u"MERGEFIELD text_Field2 \\* Upper", nullptr);
    builder->Write(u", ");
    builder->InsertField(u"MERGEFIELD numeric_Field1 \\# 0.0", nullptr);
    
    builder->get_Document()->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExMailMergeEvent::FieldValueMergingCallback>());
    
    builder->get_Document()->get_MailMerge()->Execute(System::MakeArray<System::String>({u"text_Field1", u"text_Field2", u"numeric_Field1"}), System::MakeArray<System::SharedPtr<System::Object>>({System::ExplicitCast<System::Object>(u"Field 1"), System::ExplicitCast<System::Object>(u"Field 2"), System::ExplicitCast<System::Object>(10)}));
    System::String t = doc->GetText().Trim();
    ASSERT_EQ(u"Merge Value For \"Text_Field1\": Field 1, MERGE VALUE FOR \"TEXT_FIELD2\": FIELD 2, 10000.0", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExMailMergeEvent, FieldFormats)
{
    s_instance->FieldFormats();
}

} // namespace gtest_test

void ExMailMergeEvent::ImageFromUrl()
{
    //ExStart
    //ExFor:MailMerge.Execute(String[], Object[])
    //ExSummary:Shows how to merge an image from a URI as mail merge data into a MERGEFIELD.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // MERGEFIELDs with "Image:" tags will receive an image during a mail merge.
    // The string after the colon in the "Image:" tag corresponds to a column name
    // in the data source whose cells contain URIs of image files.
    builder->InsertField(u"MERGEFIELD  Image:logo_FromWeb ");
    builder->InsertField(u"MERGEFIELD  Image:logo_FromFileSystem ");
    
    // Create a data source that contains URIs of images that we will merge. 
    // A URI can be a web URL that points to an image, or a local file system filename of an image file.
    System::ArrayPtr<System::String> columns = System::MakeArray<System::String>({u"logo_FromWeb", u"logo_FromFileSystem"});
    System::ArrayPtr<System::SharedPtr<System::Object>> URIs = System::MakeArray<System::SharedPtr<System::Object>>({System::ExplicitCast<System::Object>(get_ImageUrl()), System::ExplicitCast<System::Object>(get_ImageDir() + u"Logo.jpg")});
    
    // Execute a mail merge on a data source with one row.
    doc->get_MailMerge()->Execute(columns, URIs);
    
    doc->Save(get_ArtifactsDir() + u"MailMergeEvent.ImageFromUrl.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"MailMergeEvent.ImageFromUrl.docx");
    
    auto imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(400, 400, Aspose::Words::Drawing::ImageType::Jpeg, imageShape);
    
    imageShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    Aspose::Words::ApiExamples::TestUtil::VerifyImageInShape(272, 92, Aspose::Words::Drawing::ImageType::Png, imageShape);
}

namespace gtest_test
{

TEST_F(ExMailMergeEvent, ImageFromUrl)
{
    s_instance->ImageFromUrl();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
