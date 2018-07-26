#include "../../examples.h"

#include <system/shared_ptr.h>

#include "Model/Document/Document.h"
#include <Model/Text/Range.h>

using namespace Aspose::Words;

void RangesGetText()
{
    // ExStart:RangesGetText
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithRanges();
    auto doc = System::MakeObject<Document>(dataDir + u"Document.doc");
    System::String text = doc->get_Range()->get_Text();
    // ExEnd:RangesGetText
    std::cout << "\nDocument have following text range " << text.ToUtf8String() << std::endl;
}