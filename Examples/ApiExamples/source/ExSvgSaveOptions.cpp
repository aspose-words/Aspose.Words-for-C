// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSvgSaveOptions.h"

#include <system/object_ext.h>
#include <system/io/memory_stream.h>
#include <system/io/directory.h>
#include <system/details/dispose_guard.h>
#include <iostream>
#include <Aspose.Words.Cpp/Rendering/OfficeMathRenderer.h>
#include <Aspose.Words.Cpp/Model/Saving/SvgTextOutputMode.h>
#include <Aspose.Words.Cpp/Model/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Math;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2154726505u, ::Aspose::Words::ApiExamples::ExSvgSaveOptions::ResourceUriPrinter, ThisTypeBaseTypesInfo);

void ExSvgSaveOptions::ResourceUriPrinter::ResourceSaving(System::SharedPtr<Aspose::Words::Saving::ResourceSavingArgs> args)
{
    std::cout << System::String::Format(u"Resource #{0} \"{1}\"", ++mSavedResourceCount, args->get_ResourceFileName()) << std::endl;
    std::cout << (System::String(u"\t") + args->get_ResourceFileUri()) << std::endl;
}

ExSvgSaveOptions::ResourceUriPrinter::ResourceUriPrinter() : mSavedResourceCount(0)
{
}


RTTI_INFO_IMPL_HASH(243812173u, ::Aspose::Words::ApiExamples::ExSvgSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExSvgSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExSvgSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExSvgSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExSvgSaveOptions> ExSvgSaveOptions::s_instance;

} // namespace gtest_test

void ExSvgSaveOptions::SaveLikeImage()
{
    //ExStart
    //ExFor:SvgSaveOptions.FitToViewPort
    //ExFor:SvgSaveOptions.ShowPageBorder
    //ExFor:SvgSaveOptions.TextOutputMode
    //ExFor:SvgTextOutputMode
    //ExSummary:Shows how to mimic the properties of images when converting a .docx document to .svg.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    // Configure the SvgSaveOptions object to save with no page borders or selectable text.
    auto options = System::MakeObject<Aspose::Words::Saving::SvgSaveOptions>();
    options->set_FitToViewPort(true);
    options->set_ShowPageBorder(false);
    options->set_TextOutputMode(Aspose::Words::Saving::SvgTextOutputMode::UsePlacedGlyphs);
    
    doc->Save(get_ArtifactsDir() + u"SvgSaveOptions.SaveLikeImage.svg", options);
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
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto options = System::MakeObject<Aspose::Words::Saving::SvgSaveOptions>();
    options->set_SaveFormat(Aspose::Words::SaveFormat::Svg);
    options->set_ExportEmbeddedImages(false);
    options->set_ResourcesFolder(get_ArtifactsDir() + u"SvgResourceFolder");
    options->set_ResourcesFolderAlias(get_ArtifactsDir() + u"SvgResourceFolderAlias");
    options->set_ShowPageBorder(false);
    options->set_ResourceSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSvgSaveOptions::ResourceUriPrinter>());
    
    System::IO::Directory::CreateDirectory_(options->get_ResourcesFolderAlias());
    
    doc->Save(get_ArtifactsDir() + u"SvgSaveOptions.SvgResourceFolder.svg", options);
}

namespace gtest_test
{

TEST_F(ExSvgSaveOptions, SvgResourceFolder)
{
    s_instance->SvgResourceFolder();
}

} // namespace gtest_test

void ExSvgSaveOptions::SaveOfficeMath()
{
    //ExStart:SaveOfficeMath
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:NodeRendererBase.Save(String, SvgSaveOptions)
    //ExFor:NodeRendererBase.Save(Stream, SvgSaveOptions)
    //ExSummary:Shows how to pass save options when rendering office math.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Office math.docx");
    
    auto math = System::ExplicitCast<Aspose::Words::Math::OfficeMath>(doc->GetChild(Aspose::Words::NodeType::OfficeMath, 0, true));
    
    auto options = System::MakeObject<Aspose::Words::Saving::SvgSaveOptions>();
    options->set_TextOutputMode(Aspose::Words::Saving::SvgTextOutputMode::UsePlacedGlyphs);
    
    math->GetMathRenderer()->Save(get_ArtifactsDir() + u"SvgSaveOptions.Output.svg", options);
    
    {
        auto stream = System::MakeObject<System::IO::MemoryStream>();
        math->GetMathRenderer()->Save(stream, options);
    }
    //ExEnd:SaveOfficeMath
}

namespace gtest_test
{

TEST_F(ExSvgSaveOptions, SaveOfficeMath)
{
    s_instance->SaveOfficeMath();
}

} // namespace gtest_test

void ExSvgSaveOptions::MaxImageResolution()
{
    //ExStart:MaxImageResolution
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:ShapeBase.SoftEdge
    //ExFor:SoftEdgeFormat.Radius
    //ExFor:SoftEdgeFormat.Remove
    //ExFor:SvgSaveOptions.MaxImageResolution
    //ExSummary:Shows how to set limit for image resolution.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::SvgSaveOptions>();
    saveOptions->set_MaxImageResolution(72);
    
    doc->Save(get_ArtifactsDir() + u"SvgSaveOptions.MaxImageResolution.svg", saveOptions);
    //ExEnd:MaxImageResolution
}

namespace gtest_test
{

TEST_F(ExSvgSaveOptions, MaxImageResolution)
{
    s_instance->MaxImageResolution();
}

} // namespace gtest_test

void ExSvgSaveOptions::IdPrefixSvg()
{
    //ExStart:IdPrefixSvg
    //GistId:f86d49dc0e6781b93e576539a01e6ca2
    //ExFor:SvgSaveOptions.IdPrefix
    //ExSummary:Shows how to add a prefix that is prepended to all generated element IDs (svg).
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Id prefix.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::SvgSaveOptions>();
    saveOptions->set_IdPrefix(u"pfx1_");
    
    doc->Save(get_ArtifactsDir() + u"SvgSaveOptions.IdPrefixSvg.html", saveOptions);
    //ExEnd:IdPrefixSvg
}

namespace gtest_test
{

TEST_F(ExSvgSaveOptions, IdPrefixSvg)
{
    s_instance->IdPrefixSvg();
}

} // namespace gtest_test

void ExSvgSaveOptions::RemoveJavaScriptFromLinksSvg()
{
    //ExStart:RemoveJavaScriptFromLinksSvg
    //GistId:f86d49dc0e6781b93e576539a01e6ca2
    //ExFor:SvgSaveOptions.RemoveJavaScriptFromLinks
    //ExSummary:Shows how to remove JavaScript from the links (svg).
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"JavaScript in HREF.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::SvgSaveOptions>();
    saveOptions->set_RemoveJavaScriptFromLinks(true);
    
    doc->Save(get_ArtifactsDir() + u"SvgSaveOptions.RemoveJavaScriptFromLinksSvg.html", saveOptions);
    //ExEnd:RemoveJavaScriptFromLinksSvg
}

namespace gtest_test
{

TEST_F(ExSvgSaveOptions, RemoveJavaScriptFromLinksSvg)
{
    s_instance->RemoveJavaScriptFromLinksSvg();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
