#include "Working with PdfSaveOptions.h"

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Saving;
namespace DocsExamples { namespace File_Formats_and_Conversions { namespace Save_Options { namespace gtest_test {

class WorkingWithPdfSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithPdfSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithPdfSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::File_Formats_and_Conversions::Save_Options::WorkingWithPdfSaveOptions> WorkingWithPdfSaveOptions::s_instance;

TEST_F(WorkingWithPdfSaveOptions, DisplayDocTitleInWindowTitlebar)
{
    s_instance->DisplayDocTitleInWindowTitlebar();
}

TEST_F(WorkingWithPdfSaveOptions, PdfRenderWarnings)
{
    s_instance->PdfRenderWarnings();
}

TEST_F(WorkingWithPdfSaveOptions, DigitallySignedPdfUsingCertificateHolder)
{
    s_instance->DigitallySignedPdfUsingCertificateHolder();
}

TEST_F(WorkingWithPdfSaveOptions, EmbeddedAllFonts)
{
    s_instance->EmbeddedAllFonts();
}

TEST_F(WorkingWithPdfSaveOptions, EmbeddedSubsetFonts)
{
    s_instance->EmbeddedSubsetFonts();
}

TEST_F(WorkingWithPdfSaveOptions, DisableEmbedWindowsFonts)
{
    s_instance->DisableEmbedWindowsFonts();
}

TEST_F(WorkingWithPdfSaveOptions, SkipEmbeddedArialAndTimesRomanFonts)
{
    s_instance->SkipEmbeddedArialAndTimesRomanFonts();
}

TEST_F(WorkingWithPdfSaveOptions, AvoidEmbeddingCoreFonts)
{
    s_instance->AvoidEmbeddingCoreFonts();
}

TEST_F(WorkingWithPdfSaveOptions, ExportHeaderFooterBookmarks)
{
    s_instance->ExportHeaderFooterBookmarks();
}

TEST_F(WorkingWithPdfSaveOptions, EmulateRenderingToSizeOnPage)
{
    s_instance->EmulateRenderingToSizeOnPage();
}

TEST_F(WorkingWithPdfSaveOptions, AdditionalTextPositioning)
{
    s_instance->AdditionalTextPositioning();
}

TEST_F(WorkingWithPdfSaveOptions, ConversionToPdf17)
{
    s_instance->ConversionToPdf17();
}

TEST_F(WorkingWithPdfSaveOptions, DownsamplingImages)
{
    s_instance->DownsamplingImages();
}

TEST_F(WorkingWithPdfSaveOptions, SetOutlineOptions)
{
    s_instance->SetOutlineOptions();
}

TEST_F(WorkingWithPdfSaveOptions, CustomPropertiesExport)
{
    s_instance->CustomPropertiesExport();
}

TEST_F(WorkingWithPdfSaveOptions, ExportDocumentStructure)
{
    s_instance->ExportDocumentStructure();
}

TEST_F(WorkingWithPdfSaveOptions, ImageCompression)
{
    s_instance->ImageCompression();
}

TEST_F(WorkingWithPdfSaveOptions, UpdateLastPrintedProperty)
{
    s_instance->UpdateLastPrintedProperty();
}

TEST_F(WorkingWithPdfSaveOptions, Dml3DEffectsRendering)
{
    s_instance->Dml3DEffectsRendering();
}

TEST_F(WorkingWithPdfSaveOptions, InterpolateImages)
{
    s_instance->InterpolateImages();
}

}}}} // namespace DocsExamples::File_Formats_and_Conversions::Save_Options::gtest_test
