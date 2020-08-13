// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSvgSaveOptions.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/directory.h>
#include <system/console.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/SvgTextOutputMode.h>
#include <Aspose.Words.Cpp/Model/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/ResourceSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IResourceSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(153926191u, ::ApiExamples::ExSvgSaveOptions::ResourceUriPrinter, ThisTypeBaseTypesInfo);

void ExSvgSaveOptions::ResourceUriPrinter::ResourceSaving(SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args)
{
    // If we set a folder alias in the SaveOptions object, it will be printed here
    System::Console::WriteLine(String::Format(u"Resource #{0} \"{1}\"",++mSavedResourceCount,args->get_ResourceFileName()));
    System::Console::WriteLine(String(u"\t") + args->get_ResourceFileUri());
}

ExSvgSaveOptions::ResourceUriPrinter::ResourceUriPrinter() : mSavedResourceCount(0)
{
}

namespace gtest_test
{

class ExSvgSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExSvgSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExSvgSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExSvgSaveOptions> ExSvgSaveOptions::s_instance;

} // namespace gtest_test

void ExSvgSaveOptions::SaveLikeImage()
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
    options->set_TextOutputMode(Aspose::Words::Saving::SvgTextOutputMode::UsePlacedGlyphs);

    doc->Save(ArtifactsDir + u"SvgSaveOptions.SaveLikeImage.svg", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSvgSaveOptions, SaveLikeImage)
{
    s_instance->SaveLikeImage();
}

} // namespace gtest_test

void ExSvgSaveOptions::SvgResourceFolder()
{
    // Open a document which contains images
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto options = MakeObject<SvgSaveOptions>();
    options->set_SaveFormat(Aspose::Words::SaveFormat::Svg);
    options->set_ExportEmbeddedImages(false);
    options->set_ResourcesFolder(ArtifactsDir + u"SvgResourceFolder");
    options->set_ResourcesFolderAlias(ArtifactsDir + u"SvgResourceFolderAlias");
    options->set_ShowPageBorder(false);
    options->set_ResourceSavingCallback(MakeObject<ExSvgSaveOptions::ResourceUriPrinter>());

    System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());

    doc->Save(ArtifactsDir + u"SvgSaveOptions.SvgResourceFolder.svg", options);
}

namespace gtest_test
{

TEST_F(ExSvgSaveOptions, SvgResourceFolder)
{
    s_instance->SvgResourceFolder();
}

} // namespace gtest_test

} // namespace ApiExamples
