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
#include <Aspose.Words.Cpp/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SvgTextOutputMode.h>
#include <system/io/directory.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

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

        // Configure the SvgSaveOptions object to save with no page borders or selectable text.
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
    //ExSummary:Shows how to manipulate and print the URIs of linked resources created while converting a document to .svg.
    void SvgResourceFolder()
    {
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

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to .svg.
    /// </summary>
    class ResourceUriPrinter : public IResourceSavingCallback
    {
    public:
        void ResourceSaving(SharedPtr<ResourceSavingArgs> args) override
        {
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
