// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPclSaveOptions.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/enumerator_adapter.h>
#include <system/collections/ienumerable.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PclSaveOptions.h>
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

namespace gtest_test
{

class ExPclSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExPclSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExPclSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExPclSaveOptions> ExPclSaveOptions::s_instance;

} // namespace gtest_test

void ExPclSaveOptions::RasterizeElements()
{
    //ExStart
    //ExFor:PclSaveOptions
    //ExFor:PclSaveOptions.SaveFormat
    //ExFor:PclSaveOptions.RasterizeTransformedElements
    //ExSummary:Shows how to set whether or not to rasterize complex elements before saving.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto saveOptions = MakeObject<PclSaveOptions>();
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Pcl);
    saveOptions->set_RasterizeTransformedElements(true);

    doc->Save(ArtifactsDir + u"PclSaveOptions.RasterizeElements.pcl", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPclSaveOptions, RasterizeElements)
{
    s_instance->RasterizeElements();
}

} // namespace gtest_test

void ExPclSaveOptions::SetPrinterFont()
{
    //ExStart
    //ExFor:PclSaveOptions.AddPrinterFont(string, string)
    //ExFor:PclSaveOptions.FallbackFontName
    //ExSummary:Shows how to add information about font that is uploaded to the printer and set the font that will be used if no expected font is found in printer and built-in fonts collections.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto saveOptions = MakeObject<PclSaveOptions>();
    saveOptions->AddPrinterFont(u"Courier", u"Courier");
    saveOptions->set_FallbackFontName(u"Times New Roman");

    doc->Save(ArtifactsDir + u"PclSaveOptions.SetPrinterFont.pcl", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPclSaveOptions, SetPrinterFont)
{
    s_instance->SetPrinterFont();
}

} // namespace gtest_test

void ExPclSaveOptions::GetPreservedPaperTrayInformation()
{
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // Paper tray information is now preserved when saving document to PCL format
    // Following information is transferred from document's model to PCL file
    for (auto section : System::IterateOver(doc->get_Sections()->LINQ_OfType<SharedPtr<Section> >()))
    {
        section->get_PageSetup()->set_FirstPageTray(15);
        section->get_PageSetup()->set_OtherPagesTray(12);
    }

    doc->Save(ArtifactsDir + u"PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
}

namespace gtest_test
{

TEST_F(ExPclSaveOptions, DISABLED_GetPreservedPaperTrayInformation)
{
    s_instance->GetPreservedPaperTrayInformation();
}

} // namespace gtest_test

} // namespace ApiExamples
