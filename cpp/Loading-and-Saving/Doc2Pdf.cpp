#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/MetafileRenderingOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    // ExStart:RenderMetafileToBitmap
    class HandleDocumentWarnings : public IWarningCallback {
        using ThisType = HandleDocumentWarnings;
        using BaseType = IWarningCallback;
        using ThisTypeBaseTypesInfo = ::System::BaseTypesInfo<BaseType>;
		RTTI_INFO(ThisType, ThisTypeBaseTypesInfo);
    public:
        System::SharedPtr<WarningInfoCollection> mWarnings;
        void Warning(System::SharedPtr<WarningInfo> info) override;
        HandleDocumentWarnings();
    };

    void HandleDocumentWarnings::Warning(System::SharedPtr<WarningInfo> info)
    {
        //For now type of warnings about unsupported metafile records changed from DataLoss/UnexpectedContent to MinorFormattingLoss.
        if (info->get_WarningType() == WarningType::MinorFormattingLoss)
        {
            std::cout << "Unsupported operation: " << info->get_Description().ToUtf8String() << std::endl;
            mWarnings->Warning(info);
        }
    }

    HandleDocumentWarnings::HandleDocumentWarnings() : mWarnings(System::MakeObject<WarningInfoCollection>())
    {
    }
    
    void RenderMetafileToBitmap(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"PdfRenderWarnings.doc");

        System::SharedPtr<MetafileRenderingOptions> metafileRenderingOptions = System::MakeObject<MetafileRenderingOptions>();
        metafileRenderingOptions->set_EmulateRasterOperations(false);
        metafileRenderingOptions->set_RenderingMode(MetafileRenderingMode::VectorWithFallback);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
        System::SharedPtr<HandleDocumentWarnings> callback = System::MakeObject<HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        System::SharedPtr<PdfSaveOptions> saveOptions = System::MakeObject<PdfSaveOptions>();
        saveOptions->set_MetafileRenderingOptions(metafileRenderingOptions);

        doc->Save(outputDataDir + u"PdfSaveOptions.HandleRasterWarnings.pdf", saveOptions);
    }
    // ExEnd:RenderMetafileToBitmap

    void SaveDoc2Pdf(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:Doc2Pdf
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        System::String outputPath = outputDataDir + u"Doc2Pdf.SaveDoc2Pdf.pdf";
        // Save the document in PDF format.
        doc->Save(outputPath);
        // ExEnd:Doc2Pdf
        std::cout << "Document converted to PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void DisplayDocTitleInWindowTitlebar(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:DisplayDocTitleInWindowTitlebar
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Rendering.doc");

        System::SharedPtr<PdfSaveOptions> saveOptions = System::MakeObject<PdfSaveOptions>();
        saveOptions->set_DisplayDocTitle(true);

        System::String outputPath = outputDataDir + u"Doc2Pdf.DisplayDocTitleInWindowTitlebar.pdf";

        // Save the document in PDF format.
        doc->Save(outputPath, saveOptions);
        // ExEnd:DisplayDocTitleInWindowTitlebar
        std::cout << "Document converted to PDF successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void PdfRenderWarnings(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"PdfRenderWarnings.doc");

        // Set a SaveOptions object to not emulate raster operations.
        System::SharedPtr<PdfSaveOptions> saveOptions = System::MakeObject<PdfSaveOptions>();
        System::SharedPtr<MetafileRenderingOptions> metafileRenderingOptions = System::MakeObject<MetafileRenderingOptions>();
        metafileRenderingOptions->set_EmulateRasterOperations(false);
        metafileRenderingOptions->set_RenderingMode(MetafileRenderingMode::VectorWithFallback);
        saveOptions->set_MetafileRenderingOptions(metafileRenderingOptions);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
        System::SharedPtr<HandleDocumentWarnings> callback = System::MakeObject<HandleDocumentWarnings>();
        doc->set_WarningCallback(callback);

        doc->Save(outputDataDir + u"Doc2Pdf.PdfRenderWarnings.pdf", saveOptions);

        // While the file saves successfully, rendering warnings that occurred during saving are collected here.
        for (auto warningInfo : System::IterateOver(callback->mWarnings))
        {
            System::Console::WriteLine(warningInfo->get_Description());
        }
    }
}

void Doc2Pdf()
{
    std::cout << "Doc2Pdf example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

    SaveDoc2Pdf(inputDataDir, outputDataDir);
    DisplayDocTitleInWindowTitlebar(inputDataDir, outputDataDir);
    PdfRenderWarnings(inputDataDir, outputDataDir);
    RenderMetafileToBitmap(inputDataDir, outputDataDir);
    std::cout << "Doc2Pdf example finished." << std::endl << std::endl;
}