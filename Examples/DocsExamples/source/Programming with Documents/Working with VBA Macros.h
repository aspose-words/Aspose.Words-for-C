#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Vba/VbaModule.h>
#include <Aspose.Words.Cpp/Vba/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/Vba/VbaModuleType.h>
#include <Aspose.Words.Cpp/Vba/VbaProject.h>
#include <Aspose.Words.Cpp/Vba/VbaReference.h>
#include <Aspose.Words.Cpp/Vba/VbaReferenceCollection.h>
#include <Aspose.Words.Cpp/Vba/VbaReferenceType.h>
#include <system/array.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/linq/enumerable.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Vba;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithVba : public DocsExamplesBase
{
public:
    void CreateVbaProject()
    {
        //ExStart:CreateVbaProject
        auto doc = MakeObject<Document>();

        auto project = MakeObject<VbaProject>();
        project->set_Name(u"AsposeProject");
        doc->set_VbaProject(project);

        // Create a new module and specify a macro source code.
        auto module_ = MakeObject<VbaModule>();
        module_->set_Name(u"AsposeModule");
        module_->set_Type(VbaModuleType::ProceduralModule);
        module_->set_SourceCode(u"New source code");

        // Add module to the VBA project.
        doc->get_VbaProject()->get_Modules()->Add(module_);

        doc->Save(ArtifactsDir + u"WorkingWithVba.CreateVbaProject.docm");
        //ExEnd:CreateVbaProject
    }

    void ReadVbaMacros()
    {
        //ExStart:ReadVbaMacros
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");

        if (doc->get_VbaProject() != nullptr)
        {
            for (auto module_ : doc->get_VbaProject()->get_Modules())
            {
                std::cout << module_->get_SourceCode() << std::endl;
            }
        }
        //ExEnd:ReadVbaMacros
    }

    void ModifyVbaMacros()
    {
        //ExStart:ModifyVbaMacros
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");

        SharedPtr<VbaProject> project = doc->get_VbaProject();

        const String newSourceCode = u"Test change source code";
        project->get_Modules()->idx_get(0)->set_SourceCode(newSourceCode);

        doc->Save(ArtifactsDir + u"WorkingWithVba.ModifyVbaMacros.docm");
        //ExEnd:ModifyVbaMacros
    }

    void CloneVbaProject()
    {
        //ExStart:CloneVbaProject
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");
        auto destDoc = MakeObject<Document>();
        destDoc->set_VbaProject(doc->get_VbaProject()->Clone());

        destDoc->Save(ArtifactsDir + u"WorkingWithVba.CloneVbaProject.docm");
        //ExEnd:CloneVbaProject
    }

    void CloneVbaModule()
    {
        //ExStart:CloneVbaModule
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");
        auto destDoc = MakeObject<Document>();
        destDoc->set_VbaProject(MakeObject<VbaProject>());

        SharedPtr<VbaModule> copyModule = doc->get_VbaProject()->get_Modules()->idx_get(u"Module1")->Clone();
        destDoc->get_VbaProject()->get_Modules()->Add(copyModule);

        destDoc->Save(ArtifactsDir + u"WorkingWithVba.CloneVbaModule.docm");
        //ExEnd:CloneVbaModule
    }

    void RemoveBrokenRef()
    {
        //ExStart:RemoveReferenceFromCollectionOfReferences
        auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");

        // Find and remove the reference with some LibId path.
        const String brokenPath = u"brokenPath.dll";
        SharedPtr<VbaReferenceCollection> references = doc->get_VbaProject()->get_References();
        for (int i = references->get_Count() - 1; i >= 0; i--)
        {
            SharedPtr<VbaReference> reference = doc->get_VbaProject()->get_References()->LINQ_ElementAt(i);

            String path = GetLibIdPath(reference);
            if (path == brokenPath)
            {
                references->RemoveAt(i);
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithVba.RemoveBrokenRef.docm");
        //ExEnd:RemoveReferenceFromCollectionOfReferences
    }

protected:
    /// <summary>
    /// Returns string representing LibId path of a specified reference.
    /// </summary>
    String GetLibIdPath(SharedPtr<VbaReference> reference)
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
    /// <remarks>
    /// Please see details for the syntax at [MS-OVBA], 2.1.1.8 LibidReference.
    /// </remarks>
    String GetLibIdReferencePath(String libIdReference)
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
    /// <remarks>
    /// Please see details for the syntax at [MS-OVBA], 2.1.1.12 ProjectReference.
    /// </remarks>
    String GetLibIdProjectPath(String libIdProject)
    {
        return (libIdProject != nullptr) ? libIdProject.Substring(3) : u"";
    }
};

}} // namespace DocsExamples::Programming_with_Documents
