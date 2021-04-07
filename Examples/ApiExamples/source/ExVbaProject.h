#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModule.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModuleType.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaProject.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReference.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReferenceCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReferenceType.h>
#include <system/array.h>
#include <system/exceptions.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Vba;

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
        //ExSummary:Shows how to create a VBA project using macros.
        auto doc = MakeObject<Document>();

        // Create a new VBA project.
        auto project = MakeObject<VbaProject>();
        project->set_Name(u"Aspose.Project");
        doc->set_VbaProject(project);

        // Create a new module and specify a macro source code.
        auto module_ = MakeObject<VbaModule>();
        module_->set_Name(u"Aspose.Module");
        module_->set_Type(VbaModuleType::ProceduralModule);
        module_->set_SourceCode(u"New source code");

        // Add the module to the VBA project.
        doc->get_VbaProject()->get_Modules()->Add(module_);

        doc->Save(ArtifactsDir + u"VbaProject.CreateVBAMacros.docm");
        //ExEnd

        project = MakeObject<Document>(ArtifactsDir + u"VbaProject.CreateVBAMacros.docm")->get_VbaProject();

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
        //ExSummary:Shows how to deep clone a VBA project and module.
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");
        auto destDoc = MakeObject<Document>();

        SharedPtr<VbaProject> copyVbaProject = doc->get_VbaProject()->Clone();
        destDoc->set_VbaProject(copyVbaProject);

        // In the destination document, we already have a module named "Module1"
        // because we cloned it along with the project. We will need to remove the module.
        SharedPtr<VbaModule> oldVbaModule = destDoc->get_VbaProject()->get_Modules()->idx_get(u"Module1");
        SharedPtr<VbaModule> copyVbaModule = doc->get_VbaProject()->get_Modules()->idx_get(u"Module1")->Clone();
        destDoc->get_VbaProject()->get_Modules()->Remove(oldVbaModule);
        destDoc->get_VbaProject()->get_Modules()->Add(copyVbaModule);

        destDoc->Save(ArtifactsDir + u"VbaProject.CloneVbaProject.docm");
        //ExEnd

        SharedPtr<VbaProject> originalVbaProject = MakeObject<Document>(ArtifactsDir + u"VbaProject.CloneVbaProject.docm")->get_VbaProject();

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

    //ExStart
    //ExFor:VbaReference
    //ExFor:VbaReference.LibId
    //ExFor:VbaReferenceCollection
    //ExFor:VbaReferenceCollection.Count
    //ExFor:VbaReferenceCollection.RemoveAt(int)
    //ExFor:VbaReferenceCollection.Remove(VbaReference)
    //ExFor:VbaReferenceType
    //ExSummary:Shows how to get/remove an element from the VBA reference collection.
    void RemoveVbaReference()
    {
        const String brokenPath = u"X:\\broken.dll";
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");

        SharedPtr<VbaReferenceCollection> references = doc->get_VbaProject()->get_References();
        ASSERT_EQ(5, references->get_Count());

        for (int i = references->get_Count() - 1; i >= 0; i--)
        {
            SharedPtr<VbaReference> reference = doc->get_VbaProject()->get_References()->idx_get(i);
            String path = GetLibIdPath(reference);

            if (path == brokenPath)
            {
                references->RemoveAt(i);
            }
        }
        ASSERT_EQ(4, references->get_Count());

        references->Remove(references->idx_get(1));
        ASSERT_EQ(3, references->get_Count());

        doc->Save(ArtifactsDir + u"VbaProject.RemoveVbaReference.docm");
    }

    /// <summary>
    /// Returns string representing LibId path of a specified reference.
    /// </summary>
    static String GetLibIdPath(SharedPtr<VbaReference> reference)
    {
        switch (reference->get_Type())
        {
        case VbaReferenceType::Registered:
        case VbaReferenceType::Original:
        case VbaReferenceType::Control:
            return GetLibIdReferencePath(reference->get_LibId());

        case VbaReferenceType::Project:
            return GetLibIdProjectPath(reference->get_LibId());

        default:
            throw System::ArgumentOutOfRangeException();
        }
    }

    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    static String GetLibIdReferencePath(String libIdReference)
    {
        if (libIdReference != nullptr)
        {
            ArrayPtr<String> refParts = libIdReference.Split(MakeArray<char16_t>({u'#'}));
            if (refParts->get_Length() > 3)
            {
                return refParts[3];
            }
        }

        return u"";
    }

    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    static String GetLibIdProjectPath(String libIdProject)
    {
        return libIdProject != nullptr ? libIdProject.Substring(3) : u"";
    }
    //ExEnd
};

} // namespace ApiExamples
