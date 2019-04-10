#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/MailMerge/MailMerge.h>

using namespace Aspose::Words;

void ExecuteArray()
{
    std::cout << "ExecuteArray example started." << std::endl;
    // ExStart:ExecuteArray
    typedef System::SharedPtr<System::Object> TObjectPtr;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_MailMergeAndReporting();
    // Open an existing document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"MailMerge.ExecuteArray.doc");

    // Trim trailing and leading whitespaces mail merge values
    doc->get_MailMerge()->set_TrimWhitespaces(false);

    // Fill the fields in the document with user data.
    System::ArrayPtr<System::String> names = System::MakeArray<System::String>({u"FullName", u"Company", u"Address", u"Address2", u"City"});
    System::ArrayPtr<TObjectPtr> values = System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::String>(u"James Bond"),
                                                                         System::ObjectExt::Box<System::String>(u"MI5 Headquarters"),
                                                                         System::ObjectExt::Box<System::String>(u"Milbank"),
                                                                         System::ObjectExt::Box<System::String>(u""),
                                                                         System::ObjectExt::Box<System::String>(u"London")});
    doc->get_MailMerge()->Execute(names, values);

    System::String outputPath = dataDir + GetOutputFilePath(u"ExecuteArray.doc");
    // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
    doc->Save(outputPath);
    // ExEnd:ExecuteArray
    std::cout << "Simple Mail merge performed with array data successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ExecuteArray example finished." << std::endl << std::endl;
}