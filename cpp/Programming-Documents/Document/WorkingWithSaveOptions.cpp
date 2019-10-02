#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveMeasureUnit.h>
#include <Aspose.Words.Cpp/Model/Saving/OdtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void UpdateLastSavedTimeProperty(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:UpdateLastSavedTimeProperty
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        System::SharedPtr<OoxmlSaveOptions> options = System::MakeObject<OoxmlSaveOptions>();
        options->set_UpdateLastSavedTimeProperty(true);

        System::String outputPath = outputDataDir + u"WorkingWithSaveOptions.UpdateLastSavedTimeProperty.docx";

        // Save the document to disk.
        doc->Save(outputPath, options);
        // ExEnd:UpdateLastSavedTimeProperty
        std::cout << "Updated Last Saved Time Property successfully." << std::endl;
    }

    void SetMeasureUnitForODT(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:SetMeasureUnitForODT
        //Load the Word document
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        //Open Office uses centimeters when specifying lengths, widths and other measurable formatting
        //and content properties in documents whereas MS Office uses inches.

        System::SharedPtr<OdtSaveOptions> saveOptions = System::MakeObject<OdtSaveOptions>();
        saveOptions->set_MeasureUnit(OdtSaveMeasureUnit::Inches);

        System::String outputPath = outputDataDir + u"WorkingWithSaveOptions.SetMeasureUnitForODT.docx";

        //Save the document into ODT
        doc->Save(outputPath, saveOptions);
        // ExEnd:SetMeasureUnitForODT
        std::cout << "Set MeasureUnit for ODT successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkingWithSaveOptions()
{
    std::cout << "WorkingWithSaveOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    UpdateLastSavedTimeProperty(inputDataDir, outputDataDir);
    SetMeasureUnitForODT(inputDataDir, outputDataDir);
    std::cout << "WorkingWithSaveOptions example finished." << std::endl << std::endl;
}