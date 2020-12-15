// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Windows.Forms is not supported
#include "ExRendering.h"

#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfoCollection.h>
#include <system/shared_ptr.h>

using namespace Aspose::Words;
namespace ApiExamples { namespace gtest_test {

class ExRendering : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExRendering> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExRendering>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExRendering> ExRendering::s_instance;

TEST_F(ExRendering, SaveToPdfWithOutline)
{
    s_instance->SaveToPdfWithOutline();
}

TEST_F(ExRendering, SaveToPdfStreamOnePage)
{
    s_instance->SaveToPdfStreamOnePage();
}

TEST_F(ExRendering, SaveToPdfNoCompression)
{
    s_instance->SaveToPdfNoCompression();
}

TEST_F(ExRendering, PdfCustomOptions)
{
    s_instance->PdfCustomOptions();
}

TEST_F(ExRendering, SaveAsXps)
{
    s_instance->SaveAsXps();
}

TEST_F(ExRendering, SaveAsXpsBookFold)
{
    s_instance->SaveAsXpsBookFold();
}

TEST_F(ExRendering, SaveAsImage)
{
    s_instance->SaveAsImage();
}

TEST_F(ExRendering, SkipMono_SaveToTiffDefault)
{
    RecordProperty("category", "SkipMono");

    s_instance->SaveToTiffDefault();
}

TEST_F(ExRendering, SkipMono_SaveToTiffCompression)
{
    RecordProperty("category", "SkipMono");

    s_instance->SaveToTiffCompression();
}

TEST_F(ExRendering, SaveToImageResolution)
{
    s_instance->SaveToImageResolution();
}

TEST_F(ExRendering, SkipMono_SaveToEmf)
{
    RecordProperty("category", "SkipMono");

    s_instance->SaveToEmf();
}

TEST_F(ExRendering, SaveToImageJpegQuality)
{
    s_instance->SaveToImageJpegQuality();
}

TEST_F(ExRendering, SaveToImagePaperColor)
{
    s_instance->SaveToImagePaperColor();
}

TEST_F(ExRendering, SaveToImageStream)
{
    s_instance->SaveToImageStream();
}

TEST_F(ExRendering, RenderToSize)
{
    s_instance->RenderToSize();
}

TEST_F(ExRendering, Thumbnails)
{
    s_instance->Thumbnails();
}

TEST_F(ExRendering, UpdatePageLayout)
{
    s_instance->UpdatePageLayout();
}

TEST_F(ExRendering, SetTrueTypeFontsFolder)
{
    s_instance->SetTrueTypeFontsFolder();
}

TEST_F(ExRendering, SetFontsFoldersMultipleFolders)
{
    s_instance->SetFontsFoldersMultipleFolders();
}

TEST_F(ExRendering, SetFontsFoldersSystemAndCustomFolder)
{
    s_instance->SetFontsFoldersSystemAndCustomFolder();
}

TEST_F(ExRendering, SetSpecifyFontFolder)
{
    s_instance->SetSpecifyFontFolder();
}

TEST_F(ExRendering, SetFontSubstitutes)
{
    s_instance->SetFontSubstitutes();
}

TEST_F(ExRendering, SetSpecifyFontFolders)
{
    s_instance->SetSpecifyFontFolders();
}

TEST_F(ExRendering, AddFontSubstitutes)
{
    s_instance->AddFontSubstitutes();
}

TEST_F(ExRendering, SetDefaultFontName)
{
    s_instance->SetDefaultFontName();
}

TEST_F(ExRendering, UpdatePageLayoutWarnings)
{
    s_instance->UpdatePageLayoutWarnings();
}

TEST_F(ExRendering, EmbedFullFonts)
{
    s_instance->EmbedFullFonts();
}

TEST_F(ExRendering, SubsetFonts)
{
    s_instance->SubsetFonts();
}

TEST_F(ExRendering, DisableEmbedWindowsFonts)
{
    s_instance->DisableEmbedWindowsFonts();
}

TEST_F(ExRendering, DisableEmbedCoreFonts)
{
    s_instance->DisableEmbedCoreFonts();
}

TEST_F(ExRendering, EncryptionPermissions)
{
    s_instance->EncryptionPermissions();
}

TEST_F(ExRendering, SetNumeralFormat)
{
    s_instance->SetNumeralFormat();
}

}} // namespace ApiExamples::gtest_test
