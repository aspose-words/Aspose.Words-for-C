// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported
#include "ExMailMergeEvent.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1028058589u, ::ApiExamples::ExMailMergeEvent::HandleMergeFieldInsertHtml, ThisTypeBaseTypesInfo);

void ExMailMergeEvent::HandleMergeFieldInsertHtml::FieldMerging(SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
{
    // All merge fields that expect HTML data should be marked with some prefix, e.g. 'html'
    if (args->get_DocumentFieldName().StartsWith(u"html") && args->get_Field()->GetFieldCode().Contains(u"\\b"))
    {
        SharedPtr<FieldMergeField> field = args->get_Field();

        // Insert the text for this merge field as HTML data, using DocumentBuilder
        auto builder = MakeObject<DocumentBuilder>(args->get_Document());
        builder->MoveToMergeField(args->get_DocumentFieldName());
        builder->Write(field->get_TextBefore());
        builder->InsertHtml(System::ObjectExt::Unbox<String>(args->get_FieldValue()));

        // The HTML text itself should not be inserted
        // We have already inserted it as an HTML
        args->set_Text(u"");
    }
}

void ExMailMergeEvent::HandleMergeFieldInsertHtml::ImageFieldMerging(SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    // Do nothing
}

namespace gtest_test
{

class ExMailMergeEvent : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExMailMergeEvent> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExMailMergeEvent>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExMailMergeEvent> ExMailMergeEvent::s_instance;

} // namespace gtest_test

void ExMailMergeEvent::InsertHtml()
{
    auto doc = MakeObject<Document>(MyDir + u"Field sample - MERGEFIELD.docx");

    // Add a handler for the MergeField event
    doc->get_MailMerge()->set_FieldMergingCallback(MakeObject<ExMailMergeEvent::HandleMergeFieldInsertHtml>());

    const String html = u"<html>\r\n                    <h1>Hello world!</h1>\r\n            </html>";

    // Execute mail merge
    doc->get_MailMerge()->Execute(MakeArray<String>({u"htmlField1"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(html)}));

    // Save resulting document with a new name
    doc->Save(ArtifactsDir + u"MailMergeEvent.InsertHtml.docx");
}

namespace gtest_test
{

TEST_F(ExMailMergeEvent, InsertHtml)
{
    s_instance->InsertHtml();
}

} // namespace gtest_test

void ExMailMergeEvent::ImageFromUrl()
{
    //ExStart
    //ExFor:MailMerge.Execute(String[], Object[])
    //ExSummary:Demonstrates how to merge an image from a web address using an Image field.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertField(u"MERGEFIELD  Image:Logo ");

    // Pass a URL which points to the image to merge into the document
    doc->get_MailMerge()->Execute(MakeArray<String>({u"Logo"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(AsposeLogoUrl)}));

    doc->Save(ArtifactsDir + u"MailMergeEvent.ImageFromUrl.doc");
    //ExEnd

    // Verify the image was merged into the document
    auto logoImage = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    ASSERT_FALSE(logoImage == nullptr);
    ASSERT_TRUE(logoImage->get_HasImage());
}

namespace gtest_test
{

TEST_F(ExMailMergeEvent, ImageFromUrl)
{
    s_instance->ImageFromUrl();
}

} // namespace gtest_test

} // namespace ApiExamples
