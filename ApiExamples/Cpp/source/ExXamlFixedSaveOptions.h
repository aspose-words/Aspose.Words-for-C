#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/XamlFixedSaveOptions.h>
#include <system/array.h>
#include <system/collections/list.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
#include <system/scope_guard.h>
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
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExXamlFixedSaveOptions : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:XamlFixedSaveOptions
    //ExFor:XamlFixedSaveOptions.ResourceSavingCallback
    //ExFor:XamlFixedSaveOptions.ResourcesFolder
    //ExFor:XamlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:XamlFixedSaveOptions.SaveFormat
    //ExSummary:Shows how to print the URIs of linked resources created during conversion of a document to fixed-form .xaml.
    void ResourceFolder()
    {
        // Open a document which contains resources
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto callback = MakeObject<ExXamlFixedSaveOptions::ResourceUriPrinter>();

        auto options = MakeObject<XamlFixedSaveOptions>();
        options->set_SaveFormat(SaveFormat::XamlFixed);
        options->set_ResourcesFolder(ArtifactsDir + u"XamlFixedResourceFolder");
        options->set_ResourcesFolderAlias(ArtifactsDir + u"XamlFixedFolderAlias");
        options->set_ResourceSavingCallback(callback);

        // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
        // We must ensure the folder exists before the streams can put their resources into it
        System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

        doc->Save(ArtifactsDir + u"XamlFixedSaveOptions.ResourceFolder.xaml", options);

        for (auto resource : callback->get_Resources())
        {
            std::cout << resource << std::endl;
        }
        TestResourceFolder(callback);
        //ExSkip
    }

private:
    /// <summary>
    /// Counts and prints URIs of resources created during conversion to fixed .xaml.
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
    public:
        SharedPtr<System::Collections::Generic::List<String>> get_Resources()
        {
            return pr_Resources;
        }

        ResourceUriPrinter()
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            pr_Resources = MakeObject<System::Collections::Generic::List<String>>();
        }

        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
            // If we set a folder alias in the SaveOptions object, it will be stored here
            get_Resources()->Add(String::Format(u"Resource \"{0}\"\n\t{1}", args->get_ResourceFileName(), args->get_ResourceFileUri()));

            // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
            args->set_ResourceStream(MakeObject<System::IO::FileStream>(args->get_ResourceFileUri(), System::IO::FileMode::Create));
            args->set_KeepResourceStreamOpen(false);
        }

    private:
        SharedPtr<System::Collections::Generic::List<String>> pr_Resources;
    };
    //ExEnd

protected:
    void TestResourceFolder(SharedPtr<ExXamlFixedSaveOptions::ResourceUriPrinter> callback)
    {
        ASSERT_EQ(15, callback->get_Resources()->get_Count());
        for (auto resource : callback->get_Resources())
        {
            ASSERT_TRUE(System::IO::File::Exists(resource.Split(MakeArray<char16_t>({u'\t'}))->idx_get(1)));
        }
    }
};

} // namespace ApiExamples
