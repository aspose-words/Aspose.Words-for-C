// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSavingCallback.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/io/stream.h>
#include <system/io/path.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/directory.h>
#include <system/enum_helpers.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/XamlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SvgSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/IPageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/IImageSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/IDocumentPartSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/ICssSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentPartSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/CssStyleSheetType.h>
#include <Aspose.Words.Cpp/Model/Saving/CssSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeBase.h>
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

RTTI_INFO_IMPL_HASH(373837532u, ::ApiExamples::ExSavingCallback::CustomCssSavingCallback, ThisTypeBaseTypesInfo);

ExSavingCallback::CustomCssSavingCallback::CustomCssSavingCallback(String cssDocFilename, bool isExportNeeded, bool keepCssStreamOpen)
     : mIsExportNeeded(false), mKeepCssStreamOpen(false)
{
    mCssTextFileName = cssDocFilename;
    mIsExportNeeded = isExportNeeded;
    mKeepCssStreamOpen = keepCssStreamOpen;
}

void ExSavingCallback::CustomCssSavingCallback::CssSaving(SharedPtr<Aspose::Words::Saving::CssSavingArgs> args)
{
    // Set up the stream that will create the CSS document
    args->set_CssStream(MakeObject<System::IO::FileStream>(mCssTextFileName, System::IO::FileMode::Create));
    ASSERT_TRUE(args->get_CssStream()->get_CanWrite());
    args->set_IsExportNeeded(mIsExportNeeded);
    args->set_KeepCssStreamOpen(mKeepCssStreamOpen);

    // We can also access the original document here like this
    ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));
}

RTTI_INFO_IMPL_HASH(1209660007u, ::ApiExamples::ExSavingCallback::SavedImageRename, ThisTypeBaseTypesInfo);

ExSavingCallback::SavedImageRename::SavedImageRename(String outFileName) : mCount(0)
{
    mOutFileName = outFileName;
}

void ExSavingCallback::SavedImageRename::ImageSaving(SharedPtr<Aspose::Words::Saving::ImageSavingArgs> args)
{
    // Same filename and stream functions as above in IDocumentPartSavingCallback apply here
    String imageFileName = String::Format(u"{0} shape {1}, of type {2}{3}",mOutFileName,++mCount,args->get_CurrentShape()->get_ShapeType(),System::IO::Path::GetExtension(args->get_ImageFileName()));

    args->set_ImageFileName(imageFileName);

    args->set_ImageStream(MakeObject<System::IO::FileStream>(ArtifactsDir + imageFileName, System::IO::FileMode::Create));
    ASSERT_TRUE(args->get_ImageStream()->get_CanWrite());
    ASSERT_TRUE(args->get_IsImageAvailable());
    ASSERT_FALSE(args->get_KeepImageStreamOpen());
}

RTTI_INFO_IMPL_HASH(80125854u, ::ApiExamples::ExSavingCallback::SavedDocumentPartRename, ThisTypeBaseTypesInfo);

ExSavingCallback::SavedDocumentPartRename::SavedDocumentPartRename(String outFileName, Aspose::Words::Saving::DocumentSplitCriteria documentSplitCriteria)
     : mCount(0), mDocumentSplitCriteria(((Aspose::Words::Saving::DocumentSplitCriteria)0))
{
    mOutFileName = outFileName;
    mDocumentSplitCriteria = documentSplitCriteria;
}

void ExSavingCallback::SavedDocumentPartRename::DocumentPartSaving(SharedPtr<Aspose::Words::Saving::DocumentPartSavingArgs> args)
{
    ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));

    String partType = u"";

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

    String partFileName = String::Format(u"{0} part {1}, of type {2}{3}",mOutFileName,++mCount,partType,System::IO::Path::GetExtension(args->get_DocumentPartFileName()));

    // We can designate the filename and location of each output file either by filename
    args->set_DocumentPartFileName(partFileName);

    // Or we can make a new stream and choose the location of the file at construction
    args->set_DocumentPartStream(MakeObject<System::IO::FileStream>(ArtifactsDir + partFileName, System::IO::FileMode::Create));
    ASSERT_TRUE(args->get_DocumentPartStream()->get_CanWrite());
    ASSERT_FALSE(args->get_KeepDocumentPartStreamOpen());
}

System::Object::shared_members_type ApiExamples::ExSavingCallback::SavedDocumentPartRename::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExSavingCallback::SavedDocumentPartRename::mDocumentSplitCriteria", this->mDocumentSplitCriteria);

    return result;
}

RTTI_INFO_IMPL_HASH(2896350698u, ::ApiExamples::ExSavingCallback::CustomPageFileNamePageSavingCallback, ThisTypeBaseTypesInfo);

void ExSavingCallback::CustomPageFileNamePageSavingCallback::PageSaving(SharedPtr<Aspose::Words::Saving::PageSavingArgs> args)
{
    String outFileName = String::Format(u"{0}SavingCallback.PageFileName.Page_{1}.html",ArtifactsDir,args->get_PageIndex());

    // Specify name of the output file for the current page either in this
    args->set_PageFileName(outFileName);

    // ..or by setting up a custom stream
    args->set_PageStream(MakeObject<System::IO::FileStream>(outFileName, System::IO::FileMode::Create));
    ASSERT_FALSE(args->get_KeepPageStreamOpen());
}

namespace gtest_test
{

class ExSavingCallback : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExSavingCallback> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExSavingCallback>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExSavingCallback> ExSavingCallback::s_instance;

} // namespace gtest_test

void ExSavingCallback::CheckThatAllMethodsArePresent()
{
    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

    auto imageSaveOptions = MakeObject<ImageSaveOptions>(Aspose::Words::SaveFormat::Png);
    imageSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

    auto pdfSaveOptions = MakeObject<PdfSaveOptions>();
    pdfSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

    auto psSaveOptions = MakeObject<PsSaveOptions>();
    psSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

    auto svgSaveOptions = MakeObject<SvgSaveOptions>();
    svgSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

    auto xamlFixedSaveOptions = MakeObject<XamlFixedSaveOptions>();
    xamlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

    auto xpsSaveOptions = MakeObject<XpsSaveOptions>();
    xpsSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());
}

namespace gtest_test
{

TEST_F(ExSavingCallback, CheckThatAllMethodsArePresent)
{
    s_instance->CheckThatAllMethodsArePresent();
}

} // namespace gtest_test

void ExSavingCallback::PageFileName()
{
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    auto htmlFixedSaveOptions = MakeObject<HtmlFixedSaveOptions>();
    htmlFixedSaveOptions->set_PageIndex(0);
    htmlFixedSaveOptions->set_PageCount(doc->get_PageCount());
    htmlFixedSaveOptions->set_PageSavingCallback(MakeObject<ExSavingCallback::CustomPageFileNamePageSavingCallback>());

    doc->Save(String::Format(u"{0}SavingCallback.PageFileName.html",ArtifactsDir), htmlFixedSaveOptions);

    ArrayPtr<String> filePaths = System::IO::Directory::GetFiles(ArtifactsDir, u"SavingCallback.PageFileName.Page_*.html");

    for (int i = 0; i < doc->get_PageCount(); i++)
    {
        String file = String::Format(u"{0}SavingCallback.PageFileName.Page_{1}.html",ArtifactsDir,i);
        ASSERT_EQ(file, filePaths[i]);
        //ExSkip
    }
}

namespace gtest_test
{

TEST_F(ExSavingCallback, PageFileName)
{
    s_instance->PageFileName();
}

} // namespace gtest_test

void ExSavingCallback::DocumentParts()
{
    // Open a document to be converted to html
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
    String outFileName = u"SavingCallback.DocumentParts.Rendering.html";

    // We can use an appropriate SaveOptions subclass to customize the conversion process
    auto options = MakeObject<HtmlSaveOptions>();

    // We can use it to split a document into smaller parts, in this instance split by section breaks
    // Each part will be saved into a separate file, creating many files during the conversion process instead of just one
    options->set_DocumentSplitCriteria(Aspose::Words::Saving::DocumentSplitCriteria::SectionBreak);

    // We can set a callback to name each document part file ourselves
    options->set_DocumentPartSavingCallback(MakeObject<ExSavingCallback::SavedDocumentPartRename>(outFileName, options->get_DocumentSplitCriteria()));

    // If we convert a document that contains images into html, we will end up with one html file which links to several images
    // Each image will be in the form of a file in the local file system
    // There is also a callback that can customize the name and file system location of each image
    options->set_ImageSavingCallback(MakeObject<ExSavingCallback::SavedImageRename>(outFileName));

    // The DocumentPartSaving() and ImageSaving() methods of our callbacks will be run at this time
    doc->Save(ArtifactsDir + outFileName, options);
}

namespace gtest_test
{

TEST_F(ExSavingCallback, DocumentParts)
{
    s_instance->DocumentParts();
}

} // namespace gtest_test

void ExSavingCallback::CssSavingCallback()
{
    // Open a document to be converted to html
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // If our output document will produce a CSS stylesheet, we can use an HtmlSaveOptions to control where it is saved
    auto options = MakeObject<HtmlSaveOptions>();

    // By default, a CSS stylesheet is stored inside its HTML document, but we can have it saved to a separate file
    options->set_CssStyleSheetType(Aspose::Words::Saving::CssStyleSheetType::External);

    // We can designate a filename for our stylesheet like this
    options->set_CssStyleSheetFileName(ArtifactsDir + u"SavingCallback.CssSavingCallback.css");

    // A custom ICssSavingCallback implementation can also control where that stylesheet will be saved and linked to by the Html document
    // This callback will override the filename we specified above in options.CssStyleSheetFileName,
    // but will give us more control over the saving process
    options->set_CssSavingCallback(MakeObject<ExSavingCallback::CustomCssSavingCallback>(ArtifactsDir + u"SavingCallback.CssSavingCallback.css", true, false));

    // The CssSaving() method of our callback will be called at this stage
    doc->Save(ArtifactsDir + u"SavingCallback.CssSavingCallback.html", options);
}

namespace gtest_test
{

TEST_F(ExSavingCallback, CssSavingCallback)
{
    s_instance->CssSavingCallback();
}

} // namespace gtest_test

} // namespace ApiExamples
