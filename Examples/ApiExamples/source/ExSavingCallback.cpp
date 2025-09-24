// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSavingCallback.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/io/stream.h>
#include <system/io/path.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/XamlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3833302869u, ::Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback, ThisTypeBaseTypesInfo);

void ExSavingCallback::CustomFileNamePageSavingCallback::PageSaving(System::SharedPtr<Aspose::Words::Saving::PageSavingArgs> args)
{
    System::String outFileName = System::String::Format(u"{0}SavingCallback.PageFileNames.Page_{1}.html", get_ArtifactsDir(), args->get_PageIndex());
    
    // Below are two ways of specifying where Aspose.Words will save each page of the document.
    // 1 -  Set a filename for the output page file:
    args->set_PageFileName(outFileName);
    
    // 2 -  Create a custom stream for the output page file:
    args->set_PageStream(System::MakeObject<System::IO::FileStream>(outFileName, System::IO::FileMode::Create));
    
    ASSERT_FALSE(args->get_KeepPageStreamOpen());
}

RTTI_INFO_IMPL_HASH(160390628u, ::Aspose::Words::ApiExamples::ExSavingCallback::SavedDocumentPartRename, ThisTypeBaseTypesInfo);

ExSavingCallback::SavedDocumentPartRename::SavedDocumentPartRename(System::String outFileName, Aspose::Words::Saving::DocumentSplitCriteria documentSplitCriteria)
    : mCount(0), mDocumentSplitCriteria(((Aspose::Words::Saving::DocumentSplitCriteria)0))
{
    mOutFileName = outFileName;
    mDocumentSplitCriteria = documentSplitCriteria;
}

void ExSavingCallback::SavedDocumentPartRename::DocumentPartSaving(System::SharedPtr<Aspose::Words::Saving::DocumentPartSavingArgs> args)
{
    // We can access the entire source document via the "Document" property.
    ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));
    
    System::String partType = System::String::Empty;
    
    switch (mDocumentSplitCriteria)
    {
        case Aspose::Words::Saving::DocumentSplitCriteria::PageBreak:
            partType = u"Page";
            break;
        
        case Aspose::Words::Saving::DocumentSplitCriteria::ColumnBreak:
            partType = u"Column";
            break;
        
        case Aspose::Words::Saving::DocumentSplitCriteria::SectionBreak:
            partType = u"Section";
            break;
        
        case Aspose::Words::Saving::DocumentSplitCriteria::HeadingParagraph:
            partType = u"Paragraph from heading";
            break;
        
        default:
            break;
    }
    
    System::String partFileName = System::String::Format(u"{0} part {1}, of type {2}{3}", mOutFileName, ++mCount, partType, System::IO::Path::GetExtension(args->get_DocumentPartFileName()));
    
    // Below are two ways of specifying where Aspose.Words will save each part of the document.
    // 1 -  Set a filename for the output part file:
    args->set_DocumentPartFileName(partFileName);
    
    // 2 -  Create a custom stream for the output part file:
    args->set_DocumentPartStream(System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + partFileName, System::IO::FileMode::Create));
    
    ASSERT_TRUE(args->get_DocumentPartStream()->get_CanWrite());
    ASSERT_FALSE(args->get_KeepDocumentPartStreamOpen());
}

RTTI_INFO_IMPL_HASH(3879713185u, ::Aspose::Words::ApiExamples::ExSavingCallback::SavedImageRename, ThisTypeBaseTypesInfo);

ExSavingCallback::SavedImageRename::SavedImageRename(System::String outFileName) : mCount(0)
{
    mOutFileName = outFileName;
}

void ExSavingCallback::SavedImageRename::ImageSaving(System::SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args)
{
    System::String imageFileName = System::String::Format(u"{0} shape {1}, of type {2}{3}", mOutFileName, ++mCount, args->get_CurrentShape()->get_ShapeType(), System::IO::Path::GetExtension(args->get_ImageFileName()));
    
    // Below are two ways of specifying where Aspose.Words will save each part of the document.
    // 1 -  Set a filename for the output image file:
    args->set_ImageFileName(imageFileName);
    
    // 2 -  Create a custom stream for the output image file:
    args->set_ImageStream(System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + imageFileName, System::IO::FileMode::Create));
    
    ASSERT_TRUE(args->get_ImageStream()->get_CanWrite());
    ASSERT_TRUE(args->get_IsImageAvailable());
    ASSERT_FALSE(args->get_KeepImageStreamOpen());
}

RTTI_INFO_IMPL_HASH(454102306u, ::Aspose::Words::ApiExamples::ExSavingCallback::CustomCssSavingCallback, ThisTypeBaseTypesInfo);

ExSavingCallback::CustomCssSavingCallback::CustomCssSavingCallback(System::String cssDocFilename, bool isExportNeeded, bool keepCssStreamOpen)
    : mIsExportNeeded(false), mKeepCssStreamOpen(false)
{
    mCssTextFileName = cssDocFilename;
    mIsExportNeeded = isExportNeeded;
    mKeepCssStreamOpen = keepCssStreamOpen;
}

void ExSavingCallback::CustomCssSavingCallback::CssSaving(System::SharedPtr<Aspose::Words::Saving::CssSavingArgs> args)
{
    // We can access the entire source document via the "Document" property.
    ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));
    
    args->set_CssStream(System::MakeObject<System::IO::FileStream>(mCssTextFileName, System::IO::FileMode::Create));
    args->set_IsExportNeeded(mIsExportNeeded);
    args->set_KeepCssStreamOpen(mKeepCssStreamOpen);
    
    ASSERT_TRUE(args->get_CssStream()->get_CanWrite());
}


RTTI_INFO_IMPL_HASH(570031951u, ::Aspose::Words::ApiExamples::ExSavingCallback, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExSavingCallback : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExSavingCallback> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExSavingCallback>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExSavingCallback> ExSavingCallback::s_instance;

} // namespace gtest_test

void ExSavingCallback::CheckThatAllMethodsArePresent()
{
    auto htmlFixedSaveOptions = System::MakeObject<Aspose::Words::Saving::HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_PageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback>());
    
    auto imageSaveOptions = System::MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    imageSaveOptions->set_PageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback>());
    
    auto pdfSaveOptions = System::MakeObject<Aspose::Words::Saving::PdfSaveOptions>();
    pdfSaveOptions->set_PageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback>());
    
    auto psSaveOptions = System::MakeObject<Aspose::Words::Saving::PsSaveOptions>();
    psSaveOptions->set_PageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback>());
    
    auto svgSaveOptions = System::MakeObject<Aspose::Words::Saving::SvgSaveOptions>();
    svgSaveOptions->set_PageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback>());
    
    auto xamlFixedSaveOptions = System::MakeObject<Aspose::Words::Saving::XamlFixedSaveOptions>();
    xamlFixedSaveOptions->set_PageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback>());
    
    auto xpsSaveOptions = System::MakeObject<Aspose::Words::Saving::XpsSaveOptions>();
    xpsSaveOptions->set_PageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomFileNamePageSavingCallback>());
}

namespace gtest_test
{

TEST_F(ExSavingCallback, CheckThatAllMethodsArePresent)
{
    s_instance->CheckThatAllMethodsArePresent();
}

} // namespace gtest_test

void ExSavingCallback::DocumentPartsFileNames()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    System::String outFileName = u"SavingCallback.DocumentPartsFileNames.html";
    
    // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we convert the document to HTML.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    
    // If we save the document normally, there will be one output HTML
    // document with all the source document's contents.
    // Set the "DocumentSplitCriteria" property to "DocumentSplitCriteria.SectionBreak" to
    // save our document to multiple HTML files: one for each section.
    options->set_DocumentSplitCriteria(Aspose::Words::Saving::DocumentSplitCriteria::SectionBreak);
    
    // Assign a custom callback to the "DocumentPartSavingCallback" property to alter the document part saving logic.
    options->set_DocumentPartSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::SavedDocumentPartRename>(outFileName, options->get_DocumentSplitCriteria()));
    
    // If we convert a document that contains images into html, we will end up with one html file which links to several images.
    // Each image will be in the form of a file in the local file system.
    // There is also a callback that can customize the name and file system location of each image.
    options->set_ImageSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::SavedImageRename>(outFileName));
    
    doc->Save(get_ArtifactsDir() + outFileName, options);
}

namespace gtest_test
{

TEST_F(ExSavingCallback, DocumentPartsFileNames)
{
    s_instance->DocumentPartsFileNames();
}

} // namespace gtest_test

void ExSavingCallback::ExternalCssFilenames()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Rendering.docx");
    
    // Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we convert the document to HTML.
    auto options = System::MakeObject<Aspose::Words::Saving::HtmlSaveOptions>();
    
    // Set the "CssStylesheetType" property to "CssStyleSheetType.External" to
    // accompany a saved HTML document with an external CSS stylesheet file.
    options->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);
    
    // Below are two ways of specifying directories and filenames for output CSS stylesheets.
    // 1 -  Use the "CssStyleSheetFileName" property to assign a filename to our stylesheet:
    options->set_CssStyleSheetFileName(get_ArtifactsDir() + u"SavingCallback.ExternalCssFilenames.css");
    
    // 2 -  Use a custom callback to name our stylesheet:
    options->set_CssSavingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExSavingCallback::CustomCssSavingCallback>(get_ArtifactsDir() + u"SavingCallback.ExternalCssFilenames.css", true, false));
    
    doc->Save(get_ArtifactsDir() + u"SavingCallback.ExternalCssFilenames.html", options);
}

namespace gtest_test
{

TEST_F(ExSavingCallback, ExternalCssFilenames)
{
    s_instance->ExternalCssFilenames();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
