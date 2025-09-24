// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExVbaProject.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/exceptions.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReferenceType.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReferenceCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaProject.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModuleType.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaModule.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Vba;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3916452084u, ::Aspose::Words::ApiExamples::ExVbaProject, ThisTypeBaseTypesInfo);

System::String ExVbaProject::GetLibIdPath(System::SharedPtr<Aspose::Words::Vba::VbaReference> reference)
{
    switch (reference->get_Type())
    {
        case Aspose::Words::Vba::VbaReferenceType::Registered:
        case Aspose::Words::Vba::VbaReferenceType::Original:
        case Aspose::Words::Vba::VbaReferenceType::Control:
            return GetLibIdReferencePath(reference->get_LibId());
        
        case Aspose::Words::Vba::VbaReferenceType::Project:
            return GetLibIdProjectPath(reference->get_LibId());
        
        default: 
            throw System::ArgumentOutOfRangeException();
        
    }
}

System::String ExVbaProject::GetLibIdReferencePath(System::String libIdReference)
{
    if (libIdReference != nullptr)
    {
        System::ArrayPtr<System::String> refParts = libIdReference.Split(u'#');
        if (refParts->get_Length() > 3)
        {
            return refParts[3];
        }
    }
    
    return u"";
}

System::String ExVbaProject::GetLibIdProjectPath(System::String libIdProject)
{
    return libIdProject != nullptr ? libIdProject.Substring(3) : u"";
}


namespace gtest_test
{

class ExVbaProject : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExVbaProject> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExVbaProject>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExVbaProject> ExVbaProject::s_instance;

} // namespace gtest_test

void ExVbaProject::CreateNewVbaProject()
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a new VBA project.
    auto project = System::MakeObject<Aspose::Words::Vba::VbaProject>();
    project->set_Name(u"Aspose.Project");
    doc->set_VbaProject(project);
    
    // Create a new module and specify a macro source code.
    auto module_ = System::MakeObject<Aspose::Words::Vba::VbaModule>();
    module_->set_Name(u"Aspose.Module");
    module_->set_Type(Aspose::Words::Vba::VbaModuleType::ProceduralModule);
    module_->set_SourceCode(u"New source code");
    
    // Add the module to the VBA project.
    doc->get_VbaProject()->get_Modules()->Add(module_);
    
    doc->Save(get_ArtifactsDir() + u"VbaProject.CreateVBAMacros.docm");
    //ExEnd
    
    project = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"VbaProject.CreateVBAMacros.docm")->get_VbaProject();
    
    ASSERT_EQ(u"Aspose.Project", project->get_Name());
    
    System::SharedPtr<Aspose::Words::Vba::VbaModuleCollection> modules = doc->get_VbaProject()->get_Modules();
    
    ASSERT_EQ(2, modules->get_Count());
    
    ASSERT_EQ(u"ThisDocument", modules->idx_get(0)->get_Name());
    ASSERT_EQ(Aspose::Words::Vba::VbaModuleType::DocumentModule, modules->idx_get(0)->get_Type());
    ASSERT_TRUE(System::TestTools::IsNull(modules->idx_get(0)->get_SourceCode()));
    
    ASSERT_EQ(u"Aspose.Module", modules->idx_get(1)->get_Name());
    ASSERT_EQ(Aspose::Words::Vba::VbaModuleType::ProceduralModule, modules->idx_get(1)->get_Type());
    ASSERT_EQ(u"New source code", modules->idx_get(1)->get_SourceCode());
}

namespace gtest_test
{

TEST_F(ExVbaProject, CreateNewVbaProject)
{
    s_instance->CreateNewVbaProject();
}

} // namespace gtest_test

void ExVbaProject::CloneVbaProject()
{
    //ExStart
    //ExFor:VbaProject.Clone
    //ExFor:VbaModule.Clone
    //ExSummary:Shows how to deep clone a VBA project and module.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"VBA project.docm");
    auto destDoc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Vba::VbaProject> copyVbaProject = doc->get_VbaProject()->Clone();
    destDoc->set_VbaProject(copyVbaProject);
    
    // In the destination document, we already have a module named "Module1"
    // because we cloned it along with the project. We will need to remove the module.
    System::SharedPtr<Aspose::Words::Vba::VbaModule> oldVbaModule = destDoc->get_VbaProject()->get_Modules()->idx_get(u"Module1");
    System::SharedPtr<Aspose::Words::Vba::VbaModule> copyVbaModule = doc->get_VbaProject()->get_Modules()->idx_get(u"Module1")->Clone();
    destDoc->get_VbaProject()->get_Modules()->Remove(oldVbaModule);
    destDoc->get_VbaProject()->get_Modules()->Add(copyVbaModule);
    
    destDoc->Save(get_ArtifactsDir() + u"VbaProject.CloneVbaProject.docm");
    //ExEnd
    
    System::SharedPtr<Aspose::Words::Vba::VbaProject> originalVbaProject = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"VbaProject.CloneVbaProject.docm")->get_VbaProject();
    
    ASSERT_EQ(copyVbaProject->get_Name(), originalVbaProject->get_Name());
    ASSERT_EQ(copyVbaProject->get_CodePage(), originalVbaProject->get_CodePage());
    ASPOSE_ASSERT_EQ(copyVbaProject->get_IsSigned(), originalVbaProject->get_IsSigned());
    ASSERT_EQ(copyVbaProject->get_Modules()->get_Count(), originalVbaProject->get_Modules()->get_Count());
    
    for (int32_t i = 0; i < originalVbaProject->get_Modules()->get_Count(); i++)
    {
        ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_Name(), originalVbaProject->get_Modules()->idx_get(i)->get_Name());
        ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_Type(), originalVbaProject->get_Modules()->idx_get(i)->get_Type());
        ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_SourceCode(), originalVbaProject->get_Modules()->idx_get(i)->get_SourceCode());
    }
}

namespace gtest_test
{

TEST_F(ExVbaProject, CloneVbaProject)
{
    s_instance->CloneVbaProject();
}

} // namespace gtest_test

void ExVbaProject::RemoveVbaReference()
{
    const System::String brokenPath = u"X:\\broken.dll";
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"VBA project.docm");
    
    System::SharedPtr<Aspose::Words::Vba::VbaReferenceCollection> references = doc->get_VbaProject()->get_References();
    ASSERT_EQ(5, references->get_Count());
    
    for (int32_t i = references->get_Count() - 1; i >= 0; i--)
    {
        System::SharedPtr<Aspose::Words::Vba::VbaReference> reference = doc->get_VbaProject()->get_References()->idx_get(i);
        System::String path = GetLibIdPath(reference);
        
        if (path == brokenPath)
        {
            references->RemoveAt(i);
        }
    }
    ASSERT_EQ(4, references->get_Count());
    
    references->Remove(references->idx_get(1));
    ASSERT_EQ(3, references->get_Count());
    
    doc->Save(get_ArtifactsDir() + u"VbaProject.RemoveVbaReference.docm");
}

namespace gtest_test
{

TEST_F(ExVbaProject, RemoveVbaReference)
{
    s_instance->RemoveVbaReference();
}

} // namespace gtest_test

void ExVbaProject::IsProtected()
{
    //ExStart:IsProtected
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:VbaProject.IsProtected
    //ExSummary:Shows whether the VbaProject is password protected.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Vba protected.docm");
    ASSERT_TRUE(doc->get_VbaProject()->get_IsProtected());
    //ExEnd:IsProtected
}

namespace gtest_test
{

TEST_F(ExVbaProject, IsProtected)
{
    s_instance->IsProtected();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
