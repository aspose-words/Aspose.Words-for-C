#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageRange.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageSet.h>
#include <Aspose.Words.Cpp/Model/Saving/IPageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSavingArgs.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
        // ExStart:ConvertDocumentToPNG
        class HandlePageSavingCallback : public IPageSavingCallback
        {
            typedef HandlePageSavingCallback ThisType;
            typedef IPageSavingCallback BaseType;

            typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
            RTTI_INFO(ThisType, ThisTypeBaseTypesInfo)
        public:
            HandlePageSavingCallback(System::String const& outputDataDir)
                : mOutputDataDir(outputDataDir) {}

            void PageSaving(System::SharedPtr<Aspose::Words::Saving::PageSavingArgs> args) override
            {
                args->set_PageFileName(mOutputDataDir + System::String::Format(u"Page_{0}.png", args->get_PageIndex()));
            }

        private:
            System::String mOutputDataDir;
        };

        void ConvertDocumentToPNG(System::String const &inputDataDir, System::String const &outputDataDir)
        {
            // Load a document
            auto doc = System::MakeObject<Document>(inputDataDir + u"SampleImages.doc");

            auto imageSaveOptions = System::MakeObject<ImageSaveOptions>(SaveFormat::Png);

            auto pageRange = System::MakeObject<PageRange>(0, doc->get_PageCount() - 1);
            imageSaveOptions->set_PageSet(System::MakeObject<PageSet>( System::MakeArray<System::SharedPtr<PageRange>>({pageRange})));
            imageSaveOptions->set_PageSavingCallback(System::MakeObject<HandlePageSavingCallback>(outputDataDir));

            doc->Save(outputDataDir + u"output.png", imageSaveOptions);
            
            std::cout << "Document converted to PNG successfully.\n";
        }
        // ExEnd:ConvertDocumentToPNG
}

void ConvertWordDocument()
{

    std::cout << "ConvertWordDocument example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    ConvertDocumentToPNG(inputDataDir, outputDataDir);

    std::cout << "ConvertWordDocument example finished." << std::endl;
}
