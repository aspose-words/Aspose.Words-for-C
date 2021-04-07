#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/TaskPaneCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/WebExtensionBindingCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/WebExtensionPropertyCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/TaskPaneDockState.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionBindingType.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionStoreType.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/TaskPane.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtension.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionBinding.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionProperty.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionReference.h>
#include <system/enumerator_adapter.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::WebExtensions;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class WorkingWithWebExtension : public DocsExamplesBase
{
public:
    void UsingWebExtensionTaskPanes()
    {
        //ExStart:UsingWebExtensionTaskPanes
        auto doc = MakeObject<Document>();

        auto taskPane = MakeObject<TaskPane>();
        doc->get_WebExtensionTaskPanes()->Add(taskPane);

        taskPane->set_DockState(TaskPaneDockState::Right);
        taskPane->set_IsVisible(true);
        taskPane->set_Width(300);

        taskPane->get_WebExtension()->get_Reference()->set_Id(u"wa102923726");
        taskPane->get_WebExtension()->get_Reference()->set_Version(u"1.0.0.0");
        taskPane->get_WebExtension()->get_Reference()->set_StoreType(WebExtensionStoreType::OMEX);
        taskPane->get_WebExtension()->get_Reference()->set_Store(u"th-TH");
        taskPane->get_WebExtension()->get_Properties()->Add(MakeObject<WebExtensionProperty>(u"mailchimpCampaign", u"mailchimpCampaign"));
        taskPane->get_WebExtension()->get_Bindings()->Add(
            MakeObject<WebExtensionBinding>(u"UnnamedBinding_0_1506535429545", WebExtensionBindingType::Text, u"194740422"));

        doc->Save(ArtifactsDir + u"WorkingWithWebExtension.UsingWebExtensionTaskPanes.docx");
        //ExEnd:UsingWebExtensionTaskPanes

        //ExStart:GetListOfAddins
        doc = MakeObject<Document>(ArtifactsDir + u"WorkingWithWebExtension.UsingWebExtensionTaskPanes.docx");

        std::cout << "Task panes sources:\n" << std::endl;

        for (auto taskPaneInfo : System::IterateOver(doc->get_WebExtensionTaskPanes()))
        {
            SharedPtr<WebExtensionReference> reference = taskPaneInfo->get_WebExtension()->get_Reference();
            std::cout << "Provider: \"" << reference->get_Store() << "\", version: \"" << reference->get_Version() << "\", catalog identifier: \""
                      << reference->get_Id() << "\";" << std::endl;
        }
        //ExEnd:GetListOfAddins
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
