#pragma once
// CPPDEFECT: System.Data is not supported

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/MergeFieldImageDimension.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/MergeFieldImageDimensionUnit.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSourceRoot.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMergeCleanupOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/io/memory_stream.h>
#include <system/io/stream.h>
#include <system/object_ext.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;

namespace DocsExamples { namespace Mail_Merge_and_Reporting {

class WorkingWithFields_ : public DocsExamplesBase
{
public:
    class DataSourceRoot : public IMailMergeDataSourceRoot
    {
    private:
        class DataSource : public IMailMergeDataSource
        {
        public:
            String get_TableName() override
            {
                return TableName();
            }

            bool MoveNext() override
            {
                bool result = next;
                next = false;
                return result;
            }

            SharedPtr<IMailMergeDataSource> GetChildDataSource(String s) override
            {
                ASPOSE_UNUSED(s);
                return nullptr;
            }

            bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
            {
                ASPOSE_UNUSED(fieldName);
                fieldValue.reset();
                return false;
            }

            DataSource() : next(true)
            {
            }

        private:
            bool next;

            String TableName()
            {
                return u"example";
            }
        };

    public:
        SharedPtr<IMailMergeDataSource> GetDataSource(String s) override
        {
            ASPOSE_UNUSED(s);
            return MakeObject<WorkingWithFields_::DataSourceRoot::DataSource>();
        }
    };

    class HandleMergeImageFieldFromBlob : public IFieldMergingCallback
    {
    public:
        void FieldMerging(SharedPtr<FieldMergingArgs> args) override
        {
            ASPOSE_UNUSED(args);
            // Do nothing.
        }

        /// <summary>
        /// This is called when mail merge engine encounters Image:XXX merge field in the document.
        /// You have a chance to return an Image object, file name, or a stream that contains the image.
        /// </summary>
        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> e) override
        {
            // The field value is a byte array, just cast it and create a stream on it.
            auto imageStream = MakeObject<System::IO::MemoryStream>(System::DynamicCast<System::Array<uint8_t>>(e->get_FieldValue()));
            // Now the mail merge engine will retrieve the image from the stream.
            e->set_ImageStream(imageStream);
        }
    };

    class MailMergeSwitches : public IFieldMergingCallback
    {
    public:
        void FieldMerging(SharedPtr<FieldMergingArgs> e) override
        {
            if (e->get_FieldName().StartsWith(u"HTML"))
            {
                if (e->get_Field()->GetFieldCode().Contains(u"\\b"))
                {
                    SharedPtr<FieldMergeField> field = e->get_Field();

                    auto builder = MakeObject<DocumentBuilder>(e->get_Document());
                    builder->MoveToMergeField(e->get_DocumentFieldName(), true, false);
                    builder->Write(field->get_TextBefore());
                    builder->InsertHtml(System::ObjectExt::ToString(e->get_FieldValue()));

                    e->set_Text(u"");
                }
            }
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            ASPOSE_UNUSED(args);
        }
    };

private:
    class HandleMergeField : public IFieldMergingCallback
    {
    public:
        /// <summary>
        /// This handler is called for every mail merge field found in the document,
        /// for every record found in the data source.
        /// </summary>
        void FieldMerging(SharedPtr<FieldMergingArgs> e) override
        {
            if (mBuilder == nullptr)
            {
                mBuilder = MakeObject<DocumentBuilder>(e->get_Document());
            }

            // We decided that we want all boolean values to be output as check box form fields.
            if (System::ObjectExt::Is<bool>(e->get_FieldValue()))
            {
                // Move the "cursor" to the current merge field.
                mBuilder->MoveToMergeField(e->get_FieldName());

                String checkBoxName = String::Format(u"{0}{1}", e->get_FieldName(), e->get_RecordIndex());

                mBuilder->InsertCheckBox(checkBoxName, System::ObjectExt::Unbox<bool>(e->get_FieldValue()), 0);

                return;
            }

            if (e->get_FieldName() == u"Body")
            {
                mBuilder->MoveToMergeField(e->get_FieldName());
                mBuilder->InsertHtml(System::ObjectExt::Unbox<String>(e->get_FieldValue()));
            }
            else if (e->get_FieldName() == u"Subject")
            {
                mBuilder->MoveToMergeField(e->get_FieldName());
                String textInputName = String::Format(u"{0}{1}", e->get_FieldName(), e->get_RecordIndex());
                mBuilder->InsertTextInput(textInputName, TextFormFieldType::Regular, u"", System::ObjectExt::Unbox<String>(e->get_FieldValue()), 0);
            }
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            args->set_ImageFileName(u"Image.png");
            args->get_ImageWidth()->set_Value(200);
            args->set_ImageHeight(MakeObject<MergeFieldImageDimension>(200, MergeFieldImageDimensionUnit::Percent));
        }

    private:
        SharedPtr<DocumentBuilder> mBuilder;
    };

    class ImageFieldMergingHandler : public IFieldMergingCallback
    {
    public:
        void FieldMerging(SharedPtr<FieldMergingArgs> args) override
        {
            ASPOSE_UNUSED(args);
            //  Implementation is not required.
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            auto shape = MakeObject<Shape>(args->get_Document(), ShapeType::Image);
            shape->set_Width(126);
            shape->set_Height(126);
            shape->set_WrapType(WrapType::Square);

            shape->get_ImageData()->SetImage(MyDir + u"Mail merge image.png");

            args->set_Shape(shape);
        }
    };

public:
    void MailMergeFormFields()
    {
        //ExStart:MailMergeFormFields
        auto doc = MakeObject<Document>(MyDir + u"Mail merge destinations - Fax.docx");

        // Setup mail merge event handler to do the custom work.
        doc->get_MailMerge()->set_FieldMergingCallback(MakeObject<WorkingWithFields_::HandleMergeField>());
        // Trim trailing and leading whitespaces mail merge values.
        doc->get_MailMerge()->set_TrimWhitespaces(false);

        ArrayPtr<String> fieldNames =
            MakeArray<String>({u"RecipientName", u"SenderName", u"FaxNumber", u"PhoneNumber", u"Subject", u"Body", u"Urgent", u"ForReview", u"PleaseComment"});

        ArrayPtr<SharedPtr<System::Object>> fieldValues = MakeArray<SharedPtr<System::Object>>(
            {System::ObjectExt::Box<String>(u"Josh"), System::ObjectExt::Box<String>(u"Jenny"), System::ObjectExt::Box<String>(u"123456789"),
             System::ObjectExt::Box<String>(u""), System::ObjectExt::Box<String>(u"Hello"), System::ObjectExt::Box<String>(u"<b>HTML Body Test message 1</b>"),
             System::ObjectExt::Box<bool>(true), System::ObjectExt::Box<bool>(false), System::ObjectExt::Box<bool>(true)});

        doc->get_MailMerge()->Execute(fieldNames, fieldValues);

        doc->Save(ArtifactsDir + u"WorkingWithFields.MailMergeFormFields.docx");
        //ExEnd:MailMergeFormFields
    }

    void MailMergeImageField()
    {
        //ExStart:MailMergeImageField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"{{#foreach example}}");
        builder->Writeln(u"{{Image(126pt;126pt):stempel}}");
        builder->Writeln(u"{{/foreach example}}");

        doc->get_MailMerge()->set_UseNonMergeFields(true);
        doc->get_MailMerge()->set_TrimWhitespaces(true);
        doc->get_MailMerge()->set_UseWholeParagraphAsRegion(false);
        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyTableRows | MailMergeCleanupOptions::RemoveContainingFields |
                                                 MailMergeCleanupOptions::RemoveUnusedRegions | MailMergeCleanupOptions::RemoveUnusedFields);

        doc->get_MailMerge()->set_FieldMergingCallback(MakeObject<WorkingWithFields_::ImageFieldMergingHandler>());
        doc->get_MailMerge()->ExecuteWithRegions(MakeObject<WorkingWithFields_::DataSourceRoot>());

        doc->Save(ArtifactsDir + u"WorkingWithFields.MailMergeImageField.docx");
        //ExEnd:MailMergeImageField
    }

    void HandleMailMergeSwitches()
    {
        auto doc = MakeObject<Document>(MyDir + u"Field sample - MERGEFIELD.docx");

        doc->get_MailMerge()->set_FieldMergingCallback(MakeObject<WorkingWithFields_::MailMergeSwitches>());

        const String html = u"<html>\r\n                    <h1>Hello world!</h1>\r\n            </html>";

        doc->get_MailMerge()->Execute(MakeArray<String>({u"htmlField1"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(html)}));

        doc->Save(ArtifactsDir + u"WorkingWithFields.HandleMailMergeSwitches.docx");
    }
};

}} // namespace DocsExamples::Mail_Merge_and_Reporting
