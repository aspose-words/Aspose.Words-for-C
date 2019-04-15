#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Model/MailMerge/FieldMergingArgs.h>
#include <Model/MailMerge/IFieldMergingCallback.h>
#include <Model/MailMerge/ImageFieldMergingArgs.h>
#include <Model/MailMerge/MailMerge.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;

namespace
{
    // ExStart:HandleMailMergeSwitches
    class MailMergeSwitches : public IFieldMergingCallback
    {
        typedef MailMergeSwitches ThisType;
        typedef IFieldMergingCallback BaseType;
        typedef System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();

    public:
        void FieldMerging(System::SharedPtr<FieldMergingArgs> e) override;
        void ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args) override {}
    };

    RTTI_INFO_IMPL_HASH(273698781u, MailMergeSwitches, ThisTypeBaseTypesInfo);

    void MailMergeSwitches::FieldMerging(System::SharedPtr<FieldMergingArgs> e)
    {
        if (e->get_FieldName().StartsWith(u"HTML"))
        {
            if (e->get_Field()->GetFieldCode().Contains(u"\\b"))
            {
                System::SharedPtr<FieldMergeField> field = e->get_Field();

                System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(e->get_Document());
                builder->MoveToMergeField(e->get_DocumentFieldName(), true, false);
                builder->Write(field->get_TextBefore());
                builder->InsertHtml(System::ObjectExt::ToString(e->get_FieldValue()));

                e->set_Text(u"");
            }
        }
    }
    // ExEnd:HandleMailMergeSwitches
}

void HandleMailMergeSwitches()
{
    std::cout << "HandleMailMergeSwitches example started." << std::endl;
    typedef System::SharedPtr<System::Object> TObjectPtr;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_MailMergeAndReporting();
    // Open an existing document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"MailMergeSwitches.docx");

    doc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<MailMergeSwitches>());

    // Fill the fields in the document with user data.
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"HTML_Name"}), System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::String>(u"James Bond")}));

    System::String outputPath = dataDir + GetOutputFilePath(u"HandleMailMergeSwitches.doc");
    doc->Save(outputPath);
    std::cout << "Simple Mail merge performed with array data successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "HandleMailMergeSwitches example finished." << std::endl << std::endl;
}