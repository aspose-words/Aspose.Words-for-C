#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/TaskPane.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtension.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/TaskPaneDockState.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionReference.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionStoreType.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/TaskPaneCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionProperty.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionBinding.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionBindingType.h>

using namespace Aspose::Words;
using namespace Aspose::Words::WebExtensions;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace
{
	void UsingWebExtensionTaskPanes(System::String const &outputDataDir)
	{
		// ExStart:UsingWebExtensionTaskPanes
		System::SharedPtr<Document> doc = System::MakeObject<Document>();

		System::SharedPtr<TaskPane> taskPane = System::MakeObject<TaskPane>();
		doc->get_WebExtensionTaskPanes()->Add(taskPane);

		taskPane->set_DockState(TaskPaneDockState::Right);
		taskPane->set_IsVisible(true);
		taskPane->set_Width(300);

		taskPane->get_WebExtension()->get_Reference()->set_Id(u"wa102923726");
		taskPane->get_WebExtension()->get_Reference()->set_Version(u"1.0.0.0");
		taskPane->get_WebExtension()->get_Reference()->set_StoreType(WebExtensionStoreType::OMEX);
		taskPane->get_WebExtension()->get_Reference()->set_Store(u"th-TH");
		
		doc->Save(outputDataDir + u"output.docx", SaveFormat::Docx);
		// ExEnd:UsingWebExtensionTaskPanes 
	}
}

void WorkingWithWebExtension()
{
	std::cout << "WorkingWithWebExtensions example started." << std::endl;
	// The path to the documents directories.
	System::String inputDataDir = GetInputDataDir_WorkingWithShapes();
	System::String outputDataDir = GetOutputDataDir_WorkingWithShapes();
	UsingWebExtensionTaskPanes(outputDataDir);
	std::cout << "WorkingWithWebExtensions example finished." << std::endl << std::endl;
}