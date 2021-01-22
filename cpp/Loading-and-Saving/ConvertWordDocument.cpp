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

        void ConvertDocumentToImage(System::String const& inputDataDir, System::String const& outputDataDir)
        {
            // ExStart:ConvertDocumentToImage
            // Load the document from disk.
            auto doc = System::MakeObject<Document>(inputDataDir + u"TestDoc.docx");

            auto imageSaveOptions = System::MakeObject<ImageSaveOptions>(SaveFormat::Jpeg);

            // Set the "PageSet" to "0" to convert only the first page of a document.
            auto pageRange = System::MakeObject<PageRange>(0);
            imageSaveOptions->set_PageSet(System::MakeObject<PageSet>(System::MakeArray<System::SharedPtr<PageRange>>({ pageRange })));


            // Change the image's brightness and contrast.
            // Both are on a 0-1 scale and are at 0.5 by default.
            imageSaveOptions->set_ImageBrightness(0.3f);
            imageSaveOptions->set_ImageContrast(0.7f);

            // Change the horizontal resolution.
            // The default value for these properties is 96.0, for a resolution of 96dpi.
            imageSaveOptions->set_HorizontalResolution(72.0f);

            // Save the document in JPEG format.
            doc->Save(outputDataDir + u"SaveDocx2Jpeg.jpeg", imageSaveOptions);
            // ExEnd:ConvertDocumentToImage
        }

        void ConvertPdfToDocx(System::String const& inputDataDir, System::String const& outputDataDir)
        {
            // ExStart:ConvertPdfToDocx
            // Load the document from disk.
            auto doc = System::MakeObject<Document>(inputDataDir + u"TestDoc.pdf");

            // Save the document in DOCX format.
            doc->Save(outputDataDir + u"SavePdf2Docx.docx");
            // ExEnd:ConvertPdfToDocx
        }
}

void ConvertWordDocument()
{
    std::cout << "ConvertWordDocument example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    ConvertDocumentToPNG(inputDataDir, outputDataDir);
    ConvertDocumentToImage(inputDataDir, outputDataDir);
    ConvertPdfToDocx(inputDataDir, outputDataDir);

    std::cout << "ConvertWordDocument example finished." << std::endl;
}
