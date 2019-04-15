#include "stdafx.h"
#include "examples.h"

#include <Model/Settings/ViewType.h>
#include <Model/Settings/ViewOptions.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Settings;

void SetViewOption()
{
    std::cout << "SetViewOption example started." << std::endl;
    // ExStart:SetViewOption
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    // Load the template document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile.doc");
    // Set view option.
    doc->get_ViewOptions()->set_ViewType(ViewType::PageLayout);
    doc->get_ViewOptions()->set_ZoomPercent(50);

    System::String outputPath = dataDir + GetOutputFilePath(u"SetViewOption.doc");
    // Save the finished document.
    doc->Save(outputPath);
    // ExEnd:SetViewOption
    std::cout << "View option setup successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "SetViewOption example finished." << std::endl << std::endl;
}
