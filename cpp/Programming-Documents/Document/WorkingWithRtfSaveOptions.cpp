#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveMeasureUnit.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/RtfSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void SavingImagesAsWmf(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SavingImagesAsWmf
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");

        System::SharedPtr<RtfSaveOptions> options = System::MakeObject<RtfSaveOptions>();
        options->set_SaveImagesAsWmf(true);

        doc->Save(outputDataDir + u"WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", options);
        //ExEnd:SavingImagesAsWmf
        std::cout << "Saved document successfully with Images As Wmf." << std::endl;
    }
}

void WorkingWithRtfSaveOptions()
{
    std::cout << "WorkingWithRtfSaveOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    SavingImagesAsWmf(inputDataDir, outputDataDir);
    std::cout << "WorkingWithRtfSaveOptions example finished." << std::endl << std::endl;
}