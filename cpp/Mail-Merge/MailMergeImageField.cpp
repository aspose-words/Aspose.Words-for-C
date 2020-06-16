#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSourceRoot.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Drawing;

namespace
{
    // ExStart:ImageFieldMergingHandler
    class ImageFieldMergingHandler : public IFieldMergingCallback
    {
        typedef ImageFieldMergingHandler ThisType;
        typedef IFieldMergingCallback BaseType;
        typedef System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);

    public:
        void FieldMerging(System::SharedPtr<FieldMergingArgs> e) override {}
        void ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> args) override;
    };

    void ImageFieldMergingHandler::ImageFieldMerging(System::SharedPtr<ImageFieldMergingArgs> e)
    {
        System::SharedPtr<Shape> shape = System::MakeObject<Shape>(e->get_Document(), ShapeType::Image);
        shape->set_Width(126);
        shape->set_Height(126);
        shape->set_WrapType(WrapType::Square);

        System::String imageFileName = GetInputDataDir_MailMergeAndReporting() + u"image.png";
        shape->get_ImageData()->SetImage(imageFileName);

        e->set_Shape(shape);
    }
    // ExEnd:ImageFieldMergingHandler
    // ExStart:DataSourceRoot
    class DataSourceRoot : public IMailMergeDataSourceRoot 
    {
        public: 
            System::SharedPtr<IMailMergeDataSource> IMailMergeDataSourceRoot::GetDataSource(System::String s) override
            {
                return System::MakeObject<DataSource>();
            }

        private:
            class DataSource : public IMailMergeDataSource
            {
                private:
                    bool next { true };
                public:
                    System::String get_TableName() override {
                        return u"example";
                    }

                    bool MoveNext() override {
                        bool result = next;
                        next = false;
                        return result;
                    }
                    
                    bool GetValue(System::String fieldName, System::SharedPtr<System::Object>& fieldValue) override {
                        return false;
                    }

                    System::SharedPtr<IMailMergeDataSource> GetChildDataSource(System::String tableName) override {
                        return nullptr;
                    }
            };
    };
    // ExEnd:DataSourceRoot
}

void MailMergeImageField()
{
    std::cout << "MailMergeImageField example started." << std::endl;
    // ExStart:MailMergeImageField
    typedef System::SharedPtr<System::Object> TObjectPtr;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_MailMergeAndReporting();
    System::String outputDataDir = GetOutputDataDir_MailMergeAndReporting();
    //System::String fileName = u"Template.doc";
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"template.docx");

    doc->get_MailMerge()->set_UseNonMergeFields(true);
    doc->get_MailMerge()->set_TrimWhitespaces(true);
    doc->get_MailMerge()->set_UseWholeParagraphAsRegion(false);
    doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyTableRows
        | MailMergeCleanupOptions::RemoveContainingFields
        | MailMergeCleanupOptions::RemoveUnusedRegions
        | MailMergeCleanupOptions::RemoveUnusedFields);

    // Add a handler for the MergeField event.
    doc->get_MailMerge()->set_FieldMergingCallback(System::MakeObject<ImageFieldMergingHandler>());
    doc->get_MailMerge()->ExecuteWithRegions(System::MakeObject<DataSourceRoot>());

    System::String outputPath = outputDataDir + u"ImageMailMerge_out.doc";
    // Save the finished document.
    doc->Save(outputPath);
    // ExEnd:MailMergeImageField
    std::cout << "Mail merge performed with Image fields successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "MailMergeImageField example finished." << std::endl << std::endl;
}