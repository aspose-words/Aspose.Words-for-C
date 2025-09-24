// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExXlsxSaveOptions.h"

#include <system/string.h>
#include <Aspose.Words.Cpp/Model/Saving/XlsxSectionMode.h>
#include <Aspose.Words.Cpp/Model/Saving/XlsxSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/XlsxDateTimeParsingMode.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/CompressionLevel.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(823709592u, ::Aspose::Words::ApiExamples::ExXlsxSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExXlsxSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExXlsxSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExXlsxSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExXlsxSaveOptions> ExXlsxSaveOptions::s_instance;

} // namespace gtest_test

void ExXlsxSaveOptions::CompressXlsx()
{
    //ExStart
    //ExFor:XlsxSaveOptions
    //ExFor:XlsxSaveOptions.CompressionLevel
    //ExFor:XlsxSaveOptions.SaveFormat
    //ExSummary:Shows how to compress XLSX document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Shape with linked chart.docx");
    
    auto xlsxSaveOptions = System::MakeObject<Aspose::Words::Saving::XlsxSaveOptions>();
    xlsxSaveOptions->set_CompressionLevel(Aspose::Words::Saving::CompressionLevel::Maximum);
    xlsxSaveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Xlsx);
    
    doc->Save(get_ArtifactsDir() + u"XlsxSaveOptions.CompressXlsx.xlsx", xlsxSaveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExXlsxSaveOptions, CompressXlsx)
{
    s_instance->CompressXlsx();
}

} // namespace gtest_test

void ExXlsxSaveOptions::SelectionMode()
{
    //ExStart:SelectionMode
    //GistId:470c0da51e4317baae82ad9495747fed
    //ExFor:XlsxSaveOptions.SectionMode
    //ExFor:XlsxSectionMode
    //ExSummary:Shows how to save document as a separate worksheets.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    // Each section of a document will be created as a separate worksheet.
    // Use 'SingleWorksheet' to display all document on one worksheet.
    auto xlsxSaveOptions = System::MakeObject<Aspose::Words::Saving::XlsxSaveOptions>();
    xlsxSaveOptions->set_SectionMode(Aspose::Words::Saving::XlsxSectionMode::MultipleWorksheets);
    
    doc->Save(get_ArtifactsDir() + u"XlsxSaveOptions.SelectionMode.xlsx", xlsxSaveOptions);
    //ExEnd:SelectionMode
}

namespace gtest_test
{

TEST_F(ExXlsxSaveOptions, SelectionMode)
{
    s_instance->SelectionMode();
}

} // namespace gtest_test

void ExXlsxSaveOptions::DateTimeParsingMode()
{
    //ExStart:DateTimeParsingMode
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:XlsxSaveOptions.DateTimeParsingMode
    //ExFor:XlsxDateTimeParsingMode
    //ExSummary:Shows how to specify autodetection of the date time format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Xlsx DateTime.docx");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::XlsxSaveOptions>();
    // Specify using datetime format autodetection.
    saveOptions->set_DateTimeParsingMode(Aspose::Words::Saving::XlsxDateTimeParsingMode::Auto);
    
    doc->Save(get_ArtifactsDir() + u"XlsxSaveOptions.DateTimeParsingMode.xlsx", saveOptions);
    //ExEnd:DateTimeParsingMode
}

namespace gtest_test
{

TEST_F(ExXlsxSaveOptions, DateTimeParsingMode)
{
    s_instance->DateTimeParsingMode();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
