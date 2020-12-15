#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/RW/Ole/VbaModule.h>
#include <Aspose.Words.Cpp/RW/Ole/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/VbaModuleType.h>
#include <Aspose.Words.Cpp/RW/Ole/VbaProject.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;

namespace ApiExamples {

class ExVbaProject : public ApiExampleBase
{
public:
    void CreateNewVbaProject()
    {
        //ExStart
        //ExFor:VbaProject.#ctor
        //ExFor:VbaProject.Name
        //ExFor:VbaModule.#ctor
        //ExFor:VbaModule.Name
        //ExFor:VbaModule.Type
        //ExFor:VbaModule.SourceCode
        //ExFor:VbaModuleCollection.Add(VbaModule)
        //ExFor:VbaModuleType
        //ExSummary:Shows how to create a VbaProject from a scratch for using macros.
        auto doc = MakeObject<Document>();

        // Create a new VBA project
        auto project = MakeObject<VbaProject>();
        project->set_Name(u"Aspose.Project");
        doc->set_VbaProject(project);

        // Create a new module and specify a macro source code
        auto module_ = MakeObject<VbaModule>();
        module_->set_Name(u"Aspose.Module");
        // VbaModuleType values:
        // procedural module - A collection of subroutines and functions
        // ------
        // document module - A type of VBA project item that specifies a module for embedded macros and programmatic access
        // operations that are associated with a document
        // ------
        // class module - A module that contains the definition for a new object. Each instance of a class creates
        // a new object, and procedures that are defined in the module become properties and methods of the object
        // ------
        // designer module - A VBA module that extends the methods and properties of an ActiveX control that has been
        // registered with the project
        module_->set_Type(VbaModuleType::ProceduralModule);
        module_->set_SourceCode(u"New source code");

        // Add module to the VBA project
        doc->get_VbaProject()->get_Modules()->Add(module_);

        doc->Save(ArtifactsDir + u"Document.CreateVBAMacros.docm");
        //ExEnd

        project = MakeObject<Document>(ArtifactsDir + u"Document.CreateVBAMacros.docm")->get_VbaProject();

        ASSERT_EQ(u"Aspose.Project", project->get_Name());

        SharedPtr<VbaModuleCollection> modules = doc->get_VbaProject()->get_Modules();

        ASSERT_EQ(2, modules->get_Count());

        ASSERT_EQ(u"ThisDocument", modules->idx_get(0)->get_Name());
        ASSERT_EQ(VbaModuleType::DocumentModule, modules->idx_get(0)->get_Type());
        ASSERT_TRUE(modules->idx_get(0)->get_SourceCode() == nullptr);

        ASSERT_EQ(u"Aspose.Module", modules->idx_get(1)->get_Name());
        ASSERT_EQ(VbaModuleType::ProceduralModule, modules->idx_get(1)->get_Type());
        ASSERT_EQ(u"New source code", modules->idx_get(1)->get_SourceCode());
    }

    void CloneVbaProject()
    {
        //ExStart
        //ExFor:VbaProject.Clone
        //ExFor:VbaModule.Clone
        //ExSummary:Shows how to deep clone VbaProject and VbaModule.
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");
        auto destDoc = MakeObject<Document>();

        // Clone VbaProject to the document
        SharedPtr<VbaProject> copyVbaProject = doc->get_VbaProject()->Clone();
        destDoc->set_VbaProject(copyVbaProject);

        // In destination document we already have "Module1", because it was cloned with VbaProject
        // We will need to remove it before cloning
        SharedPtr<VbaModule> oldVbaModule = destDoc->get_VbaProject()->get_Modules()->idx_get(u"Module1");
        SharedPtr<VbaModule> copyVbaModule = doc->get_VbaProject()->get_Modules()->idx_get(u"Module1")->Clone();
        destDoc->get_VbaProject()->get_Modules()->Remove(oldVbaModule);
        destDoc->get_VbaProject()->get_Modules()->Add(copyVbaModule);

        destDoc->Save(ArtifactsDir + u"Document.CloneVbaProject.docm");
        //ExEnd

        SharedPtr<VbaProject> originalVbaProject = MakeObject<Document>(ArtifactsDir + u"Document.CloneVbaProject.docm")->get_VbaProject();

        ASSERT_EQ(copyVbaProject->get_Name(), originalVbaProject->get_Name());
        ASSERT_EQ(copyVbaProject->get_CodePage(), originalVbaProject->get_CodePage());
        ASPOSE_ASSERT_EQ(copyVbaProject->get_IsSigned(), originalVbaProject->get_IsSigned());
        ASSERT_EQ(copyVbaProject->get_Modules()->get_Count(), originalVbaProject->get_Modules()->get_Count());

        for (int i = 0; i < originalVbaProject->get_Modules()->get_Count(); i++)
        {
            ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_Name(), originalVbaProject->get_Modules()->idx_get(i)->get_Name());
            ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_Type(), originalVbaProject->get_Modules()->idx_get(i)->get_Type());
            ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_SourceCode(), originalVbaProject->get_Modules()->idx_get(i)->get_SourceCode());
        }
    }
};

} // namespace ApiExamples
