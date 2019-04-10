#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/FormFields/TextFormFieldType.h>
#include <Model/MailMerge/FieldMergingArgs.h>
#include <Model/MailMerge/IFieldMergingCallback.h>
#include <Model/MailMerge/ImageFieldMergingArgs.h>
#include <Model/MailMerge/MailMerge.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;

namespace
{
    // ExStart:HandleMergeField
    class HandleMergeField : public IFieldMergingCallback
    {
        typedef HandleMergeField ThisType;
        typedef IFieldMergingCallback BaseType;
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        void FieldMerging(System::SharedPtr<FieldMergingArgs> e) override;
        void ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args) override {}

    protected:
        System::Object::shared_members_type GetSharedMembers() override;

    private:
        System::SharedPtr<DocumentBuilder> mBuilder;
    };

    RTTI_INFO_IMPL_HASH(1055131746u, HandleMergeField, ThisTypeBaseTypesInfo);

    void HandleMergeField::FieldMerging(System::SharedPtr<FieldMergingArgs> e)
    {
        if (mBuilder == nullptr)
        {
            mBuilder = System::MakeObject<DocumentBuilder>(e->get_Document());
        }

        // We decided that we want all boolean values to be output as check box form fields.
        if (System::ObjectExt::Is<bool>(e->get_FieldValue()))
        {
            // Move the "cursor" to the current merge field.
            mBuilder->MoveToMergeField(e->get_FieldName());

            // It is nice to give names to check boxes. Lets generate a name such as MyField21 or so.
            System::String checkBoxName = System::String::Format(u"{0}{1}",e->get_FieldName(),e->get_RecordIndex());

            // Insert a check box.
            mBuilder->InsertCheckBox(checkBoxName, System::ObjectExt::Unbox<bool>(e->get_FieldValue()), 0);

            // Nothing else to do for this field.
            return;
        }

        // We want to insert html during mail merge.
        if (e->get_FieldName() == u"Body")
        {
            mBuilder->MoveToMergeField(e->get_FieldName());
            mBuilder->InsertHtml(System::ObjectExt::Unbox<System::String>(e->get_FieldValue()));
        }

        // Another example, we want the Subject field to come out as text input form field.
        if (e->get_FieldName() == u"Subject")
        {
            mBuilder->MoveToMergeField(e->get_FieldName());
            System::String textInputName = System::String::Format(u"{0}{1}",e->get_FieldName(),e->get_RecordIndex());
            mBuilder->InsertTextInput(textInputName, TextFormFieldType::Regular, u"", System::ObjectExt::Unbox<System::String>(e->get_FieldValue()), 0);
        }
    }

    System::Object::shared_members_type HandleMergeField::GetSharedMembers()
    {
        auto result = System::Object::GetSharedMembers();
        result.Add("HandleMergeField::mBuilder", this->mBuilder);
        return result;
    }
    // ExEnd:HandleMergeField
}

void MailMergeFormFields()
{
    std::cout << "MailMergeFormFields example started." << std::endl;
    // ExStart:MailMergeFormFields
    typedef System::SharedPtr<System::Object> TObjectPtr;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_MailMergeAndReporting();
    //System::String fileName = u"Template.doc";
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Template.doc");

    // Setup mail merge event handler to do the custom work.
    doc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<HandleMergeField>());

    // Trim trailing and leading whitespaces mail merge values
    doc->get_MailMerge()->set_TrimWhitespaces(false);

    // This is the data for mail merge.
    System::ArrayPtr<System::String> names = System::MakeArray<System::String>({u"RecipientName", u"SenderName", u"FaxNumber", u"PhoneNumber", u"Subject", u"Body", u"Urgent", u"ForReview", u"PleaseComment"});
    System::ArrayPtr<TObjectPtr> values = System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::String>(u"Josh"),
                                                                         System::ObjectExt::Box<System::String>(u"Jenny"),
                                                                         System::ObjectExt::Box<System::String>(u"123456789"),
                                                                         System::ObjectExt::Box<System::String>(u""),
                                                                         System::ObjectExt::Box<System::String>(u"Hello"),
                                                                         System::ObjectExt::Box<System::String>(u"<b>HTML Body Test message 1</b>"),
                                                                         System::ObjectExt::Box<bool>(true),
                                                                         System::ObjectExt::Box<bool>(false),
                                                                         System::ObjectExt::Box<bool>(true)});

    // Execute the mail merge.
    doc->get_MailMerge()->Execute(names, values);

    System::String outputPath = dataDir + GetOutputFilePath(u"MailMergeFormFields.doc");
    // Save the finished document.
    doc->Save(outputPath);
    // ExEnd:MailMergeFormFields
    std::cout << "Mail merge performed with form fields successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "MailMergeFormFields example finished." << std::endl << std::endl;
}