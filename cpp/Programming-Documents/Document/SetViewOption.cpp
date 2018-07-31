#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <Model/Settings/ViewType.h>
#include <Model/Settings/ViewOptions.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Document/Document.h>
#include <cstdint>


using namespace Aspose::Words;
using namespace Aspose::Words::Settings;

void SetViewOption()
{
    // ExStart:SetViewOption
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    // Set view option.
    doc->get_ViewOptions()->set_ViewType(Aspose::Words::Settings::ViewType::PageLayout);
    doc->get_ViewOptions()->set_ZoomPercent(50);
    
    dataDir = dataDir + GetOutputFilePath(u"SetViewOption.doc");
    // Save the finished document.
    doc->Save(dataDir);
    // ExEnd:SetViewOption
    
    std::cout << "\nView option setup successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
