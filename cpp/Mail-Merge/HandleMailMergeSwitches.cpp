#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>

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
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);

    public:
        void FieldMerging(System::SharedPtr<FieldMergingArgs> e) override;
        void ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args) override {}
    };

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
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_MailMergeAndReporting();
    System::String outputDataDir = GetOutputDataDir_MailMergeAndReporting();
    // Open an existing document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"MailMergeSwitches.docx");

    doc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<MailMergeSwitches>());

    // Fill the fields in the document with user data.
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"HTML_Name"}), System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::String>(u"James Bond")}));

    System::String outputPath = outputDataDir + u"HandleMailMergeSwitches.doc";
    doc->Save(outputPath);
    std::cout << "Simple Mail merge performed with array data successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "HandleMailMergeSwitches example finished." << std::endl << std::endl;
}