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
#include <Aspose.Words.Cpp/Model/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SvgTextOutputMode.h>
#include <system/io/directory.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExSvgSaveOptions : public ApiExampleBase
{
public:
    void SaveLikeImage()
    {
        //ExStart
        //ExFor:SvgSaveOptions.FitToViewPort
        //ExFor:SvgSaveOptions.ShowPageBorder
        //ExFor:SvgSaveOptions.TextOutputMode
        //ExFor:SvgTextOutputMode
        //ExSummary:Shows how to mimic the properties of images when converting a .docx document to .svg.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // Configure the SvgSaveOptions object to save with no page borders or selectable text
        auto options = MakeObject<SvgSaveOptions>();
        options->set_FitToViewPort(true);
        options->set_ShowPageBorder(false);
        options->set_TextOutputMode(SvgTextOutputMode::UsePlacedGlyphs);

        doc->Save(ArtifactsDir + u"SvgSaveOptions.SaveLikeImage.svg", options);
        //ExEnd
    }

    //ExStart
    //ExFor:SvgSaveOptions
    //ExFor:SvgSaveOptions.ExportEmbeddedImages
    //ExFor:SvgSaveOptions.ResourceSavingCallback
    //ExFor:SvgSaveOptions.ResourcesFolder
    //ExFor:SvgSaveOptions.ResourcesFolderAlias
    //ExFor:SvgSaveOptions.SaveFormat
    //ExSummary:Shows how to manipulate and print the URIs of linked resources created during conversion of a document to .svg.
    void SvgResourceFolder()
    {
        // Open a document which contains images
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto options = MakeObject<SvgSaveOptions>();
        options->set_SaveFormat(SaveFormat::Svg);
        options->set_ExportEmbeddedImages(false);
        options->set_ResourcesFolder(ArtifactsDir + u"SvgResourceFolder");
        options->set_ResourcesFolderAlias(ArtifactsDir + u"SvgResourceFolderAlias");
        options->set_ShowPageBorder(false);
        options->set_ResourceSavingCallback(MakeObject<ExSvgSaveOptions::ResourceUriPrinter>());

        System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

        doc->Save(ArtifactsDir + u"SvgSaveOptions.SvgResourceFolder.svg", options);
    }

private:
    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to .svg.
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
    public:
        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
            // If we set a folder alias in the SaveOptions object, it will be printed here
            std::cout << "Resource #" << ++mSavedResourceCount << " \"" << args->get_ResourceFileName() << "\"" << std::endl;
            std::cout << (String(u"\t") + args->get_ResourceFileUri()) << std::endl;
        }

        ResourceUriPrinter() : mSavedResourceCount(0)
        {
        }

    private:
        int mSavedResourceCount;
    };
    //ExEnd
};

} // namespace ApiExamples
