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
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/XamlFlowSaveOptions.h>
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

class ExXamlFlowSaveOptions : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:XamlFlowSaveOptions
    //ExFor:XamlFlowSaveOptions.#ctor
    //ExFor:XamlFlowSaveOptions.#ctor(SaveFormat)
    //ExFor:XamlFlowSaveOptions.ImageSavingCallback
    //ExFor:XamlFlowSaveOptions.ImagesFolder
    //ExFor:XamlFlowSaveOptions.ImagesFolderAlias
    //ExFor:XamlFlowSaveOptions.SaveFormat
    //ExSummary:Shows how to print the filenames of linked images created during conversion of a document to flow-form .xaml.
    void ImageFolder()
    {
        // Open a document which contains images
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto callback = MakeObject<ExXamlFlowSaveOptions::ImageUriPrinter>(ArtifactsDir + u"XamlFlowImageFolderAlias");

        auto options = MakeObject<XamlFlowSaveOptions>();
        options->set_SaveFormat(SaveFormat::XamlFlow);
        options->set_ImagesFolder(ArtifactsDir + u"XamlFlowImageFolder");
        options->set_ImagesFolderAlias(ArtifactsDir + u"XamlFlowImageFolderAlias");
        options->set_ImageSavingCallback(callback);

        // A folder specified by ImagesFolderAlias will contain the images instead of ImagesFolder
        // We must ensure the folder exists before the streams can put their images into it
        System::IO::Directory::CreateDirectory_(options->get_ImagesFolderAlias());

        doc->Save(ArtifactsDir + u"XamlFlowSaveOptions.ImageFolder.xaml", options);

        for (auto resource : callback->get_Resources())
        {
            std::cout << callback->get_ImagesFolderAlias() << "/" << resource << std::endl;
        }
        TestImageFolder(callback);
        //ExSkip
    }

private:
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
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            pr_ImagesFolderAlias = imagesFolderAlias;
            pr_Resources = MakeObject<System::Collections::Generic::List<String>>();
        }

        void ImageSaving(SharedPtr<ImageSavingArgs> args) override
        {
            get_Resources()->Add(args->get_ImageFileName());

            // If we specified a ImagesFolderAlias we will also need to redirect each stream to put its image in that folder
            args->set_ImageStream(MakeObject<System::IO::FileStream>(
                String::Format(u"{0}/{1}", get_ImagesFolderAlias(), args->get_ImageFileName()), System::IO::FileMode::Create));
            args->set_KeepImageStreamOpen(false);
        }

    private:
        String pr_ImagesFolderAlias;
        SharedPtr<System::Collections::Generic::List<String>> pr_Resources;
    };
    //ExEnd

protected:
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
