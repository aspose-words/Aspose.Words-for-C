#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>

using namespace Aspose::Words;

void RangesGetText()
{
    std::cout << "RangesGetText example started." << std::endl;
    // ExStart:RangesGetText
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithRanges();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    System::String text = doc->get_Range()->get_Text();
    // ExEnd:RangesGetText
    std::cout << "Document have following text range " << text.ToUtf8String() << std::endl;
    std::cout << "RangesGetText example finished." << std::endl << std::endl;
}