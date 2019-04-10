#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/RtfLoadOptions.h>

using namespace Aspose::Words;

void WorkingWithRTF()
{
    std::cout << "WorkingWithRTF example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_LoadingAndSaving();
    //ExStart:WorkingWithRTF
    System::SharedPtr<RtfLoadOptions> loadOptions = System::MakeObject<RtfLoadOptions>();
    loadOptions->set_RecognizeUtf8Text(true);

    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"Utf8Text.rtf", loadOptions);

    System::String outputPath = dataDir + GetOutputFilePath(u"WorkingWithRTF.rtf");
    doc->Save(outputPath);
    //ExEnd:WorkingWithRTF
    std::cout << "UTF8 text has recognized successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "WorkingWithRTF example finished." << std::endl << std::endl;
}