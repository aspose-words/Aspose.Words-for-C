// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPclSaveOptions.h"

#include <system/string.h>
#include <system/enumerator_adapter.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PclSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3396834360u, ::Aspose::Words::ApiExamples::ExPclSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExPclSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExPclSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExPclSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExPclSaveOptions> ExPclSaveOptions::s_instance;

} // namespace gtest_test

void ExPclSaveOptions::RasterizeElements()
{
    //ExStart
    //ExFor:PclSaveOptions
    //ExFor:PclSaveOptions.SaveFormat
    //ExFor:PclSaveOptions.RasterizeTransformedElements
    //ExSummary:Shows how to rasterize complex elements while saving a document to PCL.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::PclSaveOptions>();
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Pcl);
    saveOptions->set_RasterizeTransformedElements(true);
    
    doc->Save(get_ArtifactsDir() + u"PclSaveOptions.RasterizeElements.pcl", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPclSaveOptions, RasterizeElements)
{
    s_instance->RasterizeElements();
}

} // namespace gtest_test

void ExPclSaveOptions::FallbackFontName()
{
    //ExStart
    //ExFor:PclSaveOptions.FallbackFontName
    //ExSummary:Shows how to declare a font that a printer will apply to printed text as a substitute should its original font be unavailable.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Non-existent font");
    builder->Write(u"Hello world!");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::PclSaveOptions>();
    saveOptions->set_FallbackFontName(u"Times New Roman");
    
    // This document will instruct the printer to apply "Times New Roman" to the text with the missing font.
    // Should "Times New Roman" also be unavailable, the printer will default to the "Arial" font.
    doc->Save(get_ArtifactsDir() + u"PclSaveOptions.SetPrinterFont.pcl", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPclSaveOptions, FallbackFontName)
{
    s_instance->FallbackFontName();
}

} // namespace gtest_test

void ExPclSaveOptions::AddPrinterFont()
{
    //ExStart
    //ExFor:PclSaveOptions.AddPrinterFont(string, string)
    //ExSummary:Shows how to get a printer to substitute all instances of a specific font with a different font. 
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->get_Font()->set_Name(u"Courier");
    builder->Write(u"Hello world!");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::PclSaveOptions>();
    saveOptions->AddPrinterFont(u"Courier New", u"Courier");
    
    // When printing this document, the printer will use the "Courier New" font
    // to access places where our document used the "Courier" font.
    doc->Save(get_ArtifactsDir() + u"PclSaveOptions.AddPrinterFont.pcl", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPclSaveOptions, AddPrinterFont)
{
    s_instance->AddPrinterFont();
}

} // namespace gtest_test

void ExPclSaveOptions::GetPreservedPaperTrayInformation()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Paper tray information is now preserved when saving document to PCL format.
    // Following information is transferred from document's model to PCL file.
    for (auto&& section : System::IterateOver<Aspose::Words::Section>(doc->get_Sections()))
    {
        section->get_PageSetup()->set_FirstPageTray(15);
        section->get_PageSetup()->set_OtherPagesTray(12);
    }
    
    doc->Save(get_ArtifactsDir() + u"PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
}

namespace gtest_test
{

TEST_F(ExPclSaveOptions, GetPreservedPaperTrayInformation)
{
    s_instance->GetPreservedPaperTrayInformation();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
