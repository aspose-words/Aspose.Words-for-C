#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/RtfLoadOptions.h>

using namespace Aspose::Words;

void WorkingWithRTF()
{
    std::cout << "WorkingWithRTF example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    //ExStart:WorkingWithRTF
    System::SharedPtr<RtfLoadOptions> loadOptions = System::MakeObject<RtfLoadOptions>();
    loadOptions->set_RecognizeUtf8Text(true);

    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Utf8Text.rtf", loadOptions);

    System::String outputPath = outputDataDir + u"WorkingWithRTF.rtf";
    doc->Save(outputPath);
    //ExEnd:WorkingWithRTF
    std::cout << "UTF8 text has recognized successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "WorkingWithRTF example finished." << std::endl << std::endl;
}