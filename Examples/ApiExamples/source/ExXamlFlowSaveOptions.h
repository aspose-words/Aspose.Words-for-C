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
#include <Aspose.Words.Cpp/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/XamlFlowSaveOptions.h>
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

class ExXamlFlowSaveOptions : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:XamlFlowSaveOptions
    //ExFor:XamlFlowSaveOptions.#ctor()
    //ExFor:XamlFlowSaveOptions.#ctor(SaveFormat)
    //ExFor:XamlFlowSaveOptions.ImageSavingCallback
    //ExFor:XamlFlowSaveOptions.ImagesFolder
    //ExFor:XamlFlowSaveOptions.ImagesFolderAlias
    //ExFor:XamlFlowSaveOptions.SaveFormat
    //ExSummary:Shows how to print the filenames of linked images created while converting a document to flow-form .xaml.
    void ImageFolder()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto callback = MakeObject<ExXamlFlowSaveOptions::ImageUriPrinter>(ArtifactsDir + u"XamlFlowImageFolderAlias");

        // Create a "XamlFlowSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to the XAML save format.
        auto options = MakeObject<XamlFlowSaveOptions>();

        ASSERT_EQ(SaveFormat::XamlFlow, options->get_SaveFormat());

        // Use the "ImagesFolder" property to assign a folder in the local file system into which
        // Aspose.Words will save all the document's linked images.
        options->set_ImagesFolder(ArtifactsDir + u"XamlFlowImageFolder");

        // Use the "ImagesFolderAlias" property to use this folder
        // when constructing image URIs instead of the images folder's name.
        options->set_ImagesFolderAlias(ArtifactsDir + u"XamlFlowImageFolderAlias");

        options->set_ImageSavingCallback(callback);

        // A folder specified by "ImagesFolderAlias" will need to contain the resources instead of "ImagesFolder".
        // We must ensure the folder exists before the callback's streams can put their resources into it.
        System::IO::Directory::CreateDirectory_(options->get_ImagesFolderAlias());

        doc->Save(ArtifactsDir + u"XamlFlowSaveOptions.ImageFolder.xaml", options);

        for (auto resource : callback->get_Resources())
        {
            std::cout << callback->get_ImagesFolderAlias() << "/" << resource << std::endl;
        }
        TestImageFolder(callback);
        //ExSkip
    }

    /// <summary>
    /// Counts and prints filenames of images while their parent document is converted to flow-form .xaml.
    /// </summary>
    class ImageUriPrinter : public IImageSavingCallback
    {
    public:
        String get_ImagesFolderAlias()
        {
            return pr_ImagesFolderAlias;
        }

        SharedPtr<System::Collections::Generic::List<String>> get_Resources()
        {
            return pr_Resources;
        }

        ImageUriPrinter(String imagesFolderAlias)
        {
            pr_ImagesFolderAlias = imagesFolderAlias;
            pr_Resources = MakeObject<System::Collections::Generic::List<String>>();
        }

        void ImageSaving(SharedPtr<ImageSavingArgs> args) override
        {
            get_Resources()->Add(args->get_ImageFileName());

            // If we specified an image folder alias, we would also need
            // to redirect each stream to put its image in the alias folder.
            args->set_ImageStream(MakeObject<System::IO::FileStream>(String::Format(u"{0}/{1}", get_ImagesFolderAlias(), args->get_ImageFileName()),
                                                                     System::IO::FileMode::Create));
            args->set_KeepImageStreamOpen(false);
        }

    private:
        String pr_ImagesFolderAlias;
        SharedPtr<System::Collections::Generic::List<String>> pr_Resources;
    };
    //ExEnd

    void TestImageFolder(SharedPtr<ExXamlFlowSaveOptions::ImageUriPrinter> callback)
    {
        ASSERT_EQ(9, callback->get_Resources()->get_Count());
        for (auto resource : callback->get_Resources())
        {
            ASSERT_TRUE(System::IO::File::Exists(String::Format(u"{0}/{1}", callback->get_ImagesFolderAlias(), resource)));
        }
    }
};

} // namespace ApiExamples
