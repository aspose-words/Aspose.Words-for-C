#include "stdafx.h"
#include "examples.h"

#include <Model/Settings/MsWordVersion.h>
#include <Model/Settings/CompatibilityOptions.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Settings;

namespace
{
    void OptimizeFor(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:OptimizeFor
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile.docx");
        doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2016);

        System::String outputPath = outputDataDir + u"SetCompatibilityOptions.docx";

        // Save the document to disk.
        doc->Save(outputPath);
        // ExEnd:OptimizeFor
        std::cout << "Document is optimized for MS Word 2016 successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void SetCompatibilityOptions()
{
    std::cout << "SetCompatibilityOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    OptimizeFor(inputDataDir, outputDataDir);
    std::cout << "SetCompatibilityOptions example finished." << std::endl << std::endl;
}
