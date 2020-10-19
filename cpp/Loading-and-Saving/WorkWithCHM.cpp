#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>

#include <system/text/encoding.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

void WorkWithCHM()
{
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    // ExStart:LoadCHM

        
    auto options =  System::MakeObject<LoadOptions>();
    options->set_Encoding(System::Text::Encoding::GetEncoding(u"windows-1251"));
    auto doc = System::MakeObject<Document>(inputDataDir + u"help.chm", options);
    // ExEnd:LoadCHM
}
