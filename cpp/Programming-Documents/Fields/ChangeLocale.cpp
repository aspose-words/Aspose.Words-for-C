#include "stdafx.h"
#include "examples.h"

#include <system/globalization/culture_info.h>
#include <system/threading/thread.h>
#include <system/date_time.h>
#include <system/object.h>
#include <system/object_ext.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/MailMerge/MailMerge.h>

using namespace Aspose::Words;

void ChangeLocale()
{
    std::cout << "ChangeLocale example started." << std::endl;
    // ExStart:ChangeLocale
    typedef System::SharedPtr<System::Object> TObjectPtr;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();

    // Create a blank document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> b = System::MakeObject<DocumentBuilder>(doc);
    b->InsertField(u"MERGEFIELD Date");

    // Store the current culture so it can be set back once mail merge is complete.
    System::SharedPtr<System::Globalization::CultureInfo> currentCulture = System::Threading::Thread::get_CurrentThread()->get_CurrentCulture();
    // Set to German language so dates and numbers are formatted using this culture during mail merge.
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::MakeObject<System::Globalization::CultureInfo>(u"de-DE"));

    // Execute mail merge.
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"Date"}), System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::DateTime>(System::DateTime::get_Now())}));

    // Restore the original culture.
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(currentCulture);
    System::String outputPath = dataDir + GetOutputFilePath(u"ChangeLocale.doc");
    doc->Save(outputPath);
    // ExEnd:ChangeLocale

    std::cout << "Culture changed successfully used in formatting fields during update." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ChangeLocale example finished." << std::endl << std::endl;
}