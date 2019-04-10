#include "stdafx.h"
#include "examples.h"

#include "Model/Document/Document.h"
#include <Model/Text/Range.h>

using namespace Aspose::Words;

void RangesGetText()
{
    std::cout << "RangesGetText example started." << std::endl;
    // ExStart:RangesGetText
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithRanges();
    auto doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    System::String text = doc->get_Range()->get_Text();
    // ExEnd:RangesGetText
    std::cout << "Document have following text range " << text.ToUtf8String() << std::endl;
    std::cout << "RangesGetText example finished." << std::endl << std::endl;
}