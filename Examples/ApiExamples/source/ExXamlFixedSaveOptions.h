#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/IResourceSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/XamlFixedSaveOptions.h>
#include <system/array.h>
#include <system/collections/list.h>
#include <system/enumerator_adapter.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/io/stream.h>
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
    //ExSummary:Shows how to print the URIs of linked resources created while converting a document to fixed-form .xaml.
    void ResourceFolder()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
        auto callback = MakeObject<ExXamlFixedSaveOptions::ResourceUriPrinter>();

        // Create a "XamlFixedSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to the XAML save format.
        auto options = MakeObject<XamlFixedSaveOptions>();

        ASSERT_EQ(SaveFormat::XamlFixed, options->get_SaveFormat());

        // Use the "ResourcesFolder" property to assign a folder in the local file system into which
        // Aspose.Words will save all the document's linked resources, such as images and fonts.
        options->set_ResourcesFolder(ArtifactsDir + u"XamlFixedResourceFolder");

        // Use the "ResourcesFolderAlias" property to use this folder
        // when constructing image URIs instead of the resources folder's name.
        options->set_ResourcesFolderAlias(ArtifactsDir + u"XamlFixedFolderAlias");

        options->set_ResourceSavingCallback(callback);

        // A folder specified by "ResourcesFolderAlias" will need to contain the resources instead of "ResourcesFolder".
        // We must ensure the folder exists before the callback's streams can put their resources into it.
        System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

        doc->Save(ArtifactsDir + u"XamlFixedSaveOptions.ResourceFolder.xaml", options);

        for (const auto& resource : callback->get_Resources())
        {
            std::cout << resource << std::endl;
        }
        TestResourceFolder(callback);
        //ExSkip
    }

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
            pr_Resources = MakeObject<System::Collections::Generic::List<String>>();
        }

        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
            get_Resources()->Add(String::Format(u"Resource \"{0}\"\n\t{1}", args->get_ResourceFileName(), args->get_ResourceFileUri()));

            // If we specified a resource folder alias, we would also need
            // to redirect each stream to put its resource in the alias folder.
            args->set_ResourceStream(MakeObject<System::IO::FileStream>(args->get_ResourceFileUri(), System::IO::FileMode::Create));
            args->set_KeepResourceStreamOpen(false);
        }

    private:
        SharedPtr<System::Collections::Generic::List<String>> pr_Resources;
    };
    //ExEnd

    void TestResourceFolder(SharedPtr<ExXamlFixedSaveOptions::ResourceUriPrinter> callback)
    {
        ASSERT_EQ(15, callback->get_Resources()->get_Count());
        for (const auto& resource : callback->get_Resources())
        {
            ASSERT_TRUE(System::IO::File::Exists(resource.Split(MakeArray<char16_t>({u'\t'}))->idx_get(1)));
        }
    }
};

} // namespace ApiExamples
