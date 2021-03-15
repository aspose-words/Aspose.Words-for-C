#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/TxtLeadingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Document/TxtLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/TxtTrailingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtListIndentation.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModule.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaProject.h>


using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void CreateVbaMacros(System::String const &outputDataDir)
    {
        //ExStart:CreateVbaProject
        System::SharedPtr<Document> doc = System::MakeObject<Document>();

        // Create a new VBA project.
        System::SharedPtr<Vba::VbaProject> project = System::MakeObject<Vba::VbaProject>();
        project->set_Name(u"AsposeProject");
        doc->set_VbaProject(project);

        // Create a new module and specify a macro source code.
        System::SharedPtr<Vba::VbaModule> vbModule = System::MakeObject<Vba::VbaModule>();
        vbModule->set_Name(u"AsposeModule");
        vbModule->set_Type(Vba::VbaModuleType::ProceduralModule);
        vbModule->set_SourceCode(u"New source code");

        // Add module to the VBA project.
        doc->get_VbaProject()->get_Modules()->Add(vbModule);

        doc->Save(outputDataDir +  u"WorkingWithVbaMacros.CreateVbaMacros.docm");
        //ExEnd:CreateVbaProject
    }

    void ReadVbaMacros(System::String const &inputDataDir)
    {
        //ExStart:ReadVbaMacros
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"VbaProject.docm");

        if (doc->get_VbaProject() != nullptr)
        {
            for (System::SharedPtr<Vba::VbaModule> module : System::IterateOver(doc->get_VbaProject()->get_Modules()))
            {
                std::cout << module->get_SourceCode().ToUtf8String() << std::endl;
            }
        }
        //ExEnd:ReadVbaMacros
    }

    void ModifyVbaMacros(System::String const &inputDataDir)
    {
        //ExStart:ModifyVbaMacros
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"VbaProject.docm");
        System::SharedPtr<Vba::VbaProject> project = doc->get_VbaProject();

        System::String newSourceCode = u"Test change source code";

        // Choose a module, and set a new source code.
        project->get_Modules()->idx_get(0)->set_SourceCode(newSourceCode);
        //ExEnd:ModifyVbaMacros
    }

    void CloneVbaProject(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:CloneVbaProject
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"VbaProject.docm");
        System::SharedPtr<Vba::VbaProject> project = doc->get_VbaProject();

        System::SharedPtr<Document> destDoc = System::MakeObject<Document>();

        // Clone the whole project.
        destDoc->set_VbaProject(doc->get_VbaProject()->Clone());

        destDoc->Save(outputDataDir + u"WorkingWithVbaMacros.CloneVbaProject.docm");
        //ExEnd:CloneVbaProject
    }

    void CloneVbaModule(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        //ExStart:CloneVbaModule
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"VbaProject.docm");
        System::SharedPtr<Vba::VbaProject> project = doc->get_VbaProject();

        System::SharedPtr<Document> destDoc = System::MakeObject<Document>();

        destDoc->set_VbaProject(System::MakeObject<Vba::VbaProject>());

        // Clone a single module.
        System::SharedPtr<Vba::VbaModule> copyModule = doc->get_VbaProject()->get_Modules()->idx_get(u"AsposeModule")->Clone();
        destDoc->get_VbaProject()->get_Modules()->Add(copyModule);

        destDoc->Save(outputDataDir + u"WorkingWithVbaMacros.CloneVbaModule.docm");
        //ExEnd:CloneVbaModule
    }
}

void WorkingWithVbaMacros()
{
    std::cout << "WorkingWithVbaMacros example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    CreateVbaMacros(outputDataDir);
    ReadVbaMacros(inputDataDir);
    ModifyVbaMacros(inputDataDir);
    CloneVbaProject(inputDataDir, outputDataDir);
    CloneVbaModule(inputDataDir, outputDataDir);
    std::cout << "WorkingWithVbaMacros example finished." << std::endl << std::endl;
}