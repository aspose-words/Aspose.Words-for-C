#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/MailMerge/MailMerge.h>

using namespace Aspose::Words;

void MailMergeAndConditionalField()
{
    std::cout << "MailMergeAndConditionalField example started." << std::endl;
    // ExStart:MailMergeAndConditionalField
    typedef System::SharedPtr<System::Object> TObjectPtr;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_MailMergeAndReporting();
    // Open an existing document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"UnconditionalMergeFieldsAndRegions.docx");

    //Merge fields and merge regions are merged regardless of the parent IF field's condition.
    doc->get_MailMerge()->set_UnconditionalMergeFieldsAndRegions(true);

    // Fill the fields in the document with user data.
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"FullName"}), System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::String>(u"James Bond")}));

    System::String outputPath = dataDir + GetOutputFilePath(u"MailMergeAndConditionalField.docx");
    doc->Save(outputPath);
    // ExEnd:MailMergeAndConditionalField
    std::cout << "Mail merge with conditional field has performed successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "MailMergeAndConditionalField example finished." << std::endl << std::endl;
}