// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOoxmlSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/random.h>
#include <system/object_ext.h>
#include <system/nullable.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/exceptions.h>
#include <system/diagnostics/stopwatch.h>
#include <system/details/dispose_guard.h>
#include <system/default.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/imaging/image_format.h>
#include <drawing/graphics.h>
#include <drawing/color.h>
#include <drawing/bitmap.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/Zip64Mode.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/DigitalSignatureDetails.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>

#include "TestUtil.h"


using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3782713318u, ::Aspose::Words::ApiExamples::ExOoxmlSaveOptions::SavingProgressCallback, ThisTypeBaseTypesInfo);

const double ExOoxmlSaveOptions::SavingProgressCallback::MaxDuration = 0.01;

ExOoxmlSaveOptions::SavingProgressCallback::SavingProgressCallback()
{
    mSavingStartedAt = System::DateTime::get_Now();
}

void ExOoxmlSaveOptions::SavingProgressCallback::Notify(System::SharedPtr<Aspose::Words::Saving::DocumentSavingArgs> args)
{
    System::DateTime canceledAt = System::DateTime::get_Now();
    double ellapsedSeconds = (canceledAt - mSavingStartedAt).get_TotalSeconds();
    if (ellapsedSeconds > MaxDuration)
    {
        throw System::OperationCanceledException(System::String::Format(u"EstimatedProgress = {0}; CanceledAt = {1}", args->get_EstimatedProgress(), canceledAt));
    }
}


RTTI_INFO_IMPL_HASH(1452966906u, ::Aspose::Words::ApiExamples::ExOoxmlSaveOptions, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExOoxmlSaveOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExOoxmlSaveOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExOoxmlSaveOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExOoxmlSaveOptions> ExOoxmlSaveOptions::s_instance;

} // namespace gtest_test

void ExOoxmlSaveOptions::Password()
{
    //ExStart
    //ExFor:OoxmlSaveOptions.Password
    //ExSummary:Shows how to create a password encrypted Office Open XML document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Password(u"MyPassword");
    
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.Password.docx", saveOptions);
    
    // We will not be able to open this document with Microsoft Word or
    // Aspose.Words without providing the correct password.
    ASSERT_THROW(static_cast<std::function<void()>>([&doc]() -> void
    {
        doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OoxmlSaveOptions.Password.docx");
    })(), Aspose::Words::IncorrectPasswordException);
    
    // Open the encrypted document by passing the correct password in a LoadOptions object.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OoxmlSaveOptions.Password.docx", System::MakeObject<Aspose::Words::Loading::LoadOptions>(u"MyPassword"));
    
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, Password)
{
    s_instance->Password();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::Iso29500Strict()
{
    //ExStart
    //ExFor:CompatibilityOptions
    //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
    //ExFor:OoxmlSaveOptions
    //ExFor:OoxmlSaveOptions.#ctor
    //ExFor:OoxmlSaveOptions.SaveFormat
    //ExFor:OoxmlCompliance
    //ExFor:OoxmlSaveOptions.Compliance
    //ExFor:ShapeMarkupLanguage
    //ExSummary:Shows how to set an OOXML compliance specification for a saved document to adhere to.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // If we configure compatibility options to comply with Microsoft Word 2003,
    // inserting an image will define its shape using VML.
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2003);
    builder->InsertImage(get_ImageDir() + u"Transparent background logo.png");
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeMarkupLanguage::Vml, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_MarkupLanguage());
    
    // The "ISO/IEC 29500:2008" OOXML standard does not support VML shapes.
    // If we set the "Compliance" property of the SaveOptions object to "OoxmlCompliance.Iso29500_2008_Strict",
    // any document we save while passing this object will have to follow that standard. 
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Strict);
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Docx);
    
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);
    
    // Our saved document defines the shape using DML to adhere to the "ISO/IEC 29500:2008" OOXML standard.
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OoxmlSaveOptions.Iso29500Strict.docx");
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeMarkupLanguage::Dml, (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_MarkupLanguage());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, Iso29500Strict)
{
    s_instance->Iso29500Strict();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::RestartingDocumentList(bool restartListAtEachSection)
{
    //ExStart
    //ExFor:List.IsRestartAtEachSection
    //ExFor:OoxmlCompliance
    //ExFor:OoxmlSaveOptions.Compliance
    //ExSummary:Shows how to configure a list to restart numbering at each section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);
    
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->idx_get(0);
    list->set_IsRestartAtEachSection(restartListAtEachSection);
    
    // The "IsRestartAtEachSection" property will only be applicable when
    // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
    auto options = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    options->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Transitional);
    
    builder->get_ListFormat()->set_List(list);
    
    builder->Writeln(u"List item 1");
    builder->Writeln(u"List item 2");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Writeln(u"List item 3");
    builder->Writeln(u"List item 4");
    
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.RestartingDocumentList.docx", options);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OoxmlSaveOptions.RestartingDocumentList.docx");
    
    ASPOSE_ASSERT_EQ(restartListAtEachSection, doc->get_Lists()->idx_get(0)->get_IsRestartAtEachSection());
    //ExEnd
}

namespace gtest_test
{

using ExOoxmlSaveOptions_RestartingDocumentList_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExOoxmlSaveOptions::RestartingDocumentList)>::type;

struct ExOoxmlSaveOptions_RestartingDocumentList : public ExOoxmlSaveOptions, public Aspose::Words::ApiExamples::ExOoxmlSaveOptions, public ::testing::WithParamInterface<ExOoxmlSaveOptions_RestartingDocumentList_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_RestartingDocumentList, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RestartingDocumentList(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_RestartingDocumentList, ::testing::ValuesIn(ExOoxmlSaveOptions_RestartingDocumentList::TestCases()));

} // namespace gtest_test

void ExOoxmlSaveOptions::LastSavedTime(bool updateLastSavedTimeProperty)
{
    //ExStart
    //ExFor:SaveOptions.UpdateLastSavedTimeProperty
    //ExSummary:Shows how to determine whether to preserve the document's "Last saved time" property when saving.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    ASSERT_EQ(System::DateTime(2021, 5, 11, 6, 32, 0), doc->get_BuiltInDocumentProperties()->get_LastSavedTime());
    
    // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
    // and then pass it to the document's saving method to modify how we save the document.
    // Set the "UpdateLastSavedTimeProperty" property to "true" to
    // set the output document's "Last saved time" built-in property to the current date/time.
    // Set the "UpdateLastSavedTimeProperty" property to "false" to
    // preserve the original value of the input document's "Last saved time" built-in property.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_UpdateLastSavedTimeProperty(updateLastSavedTimeProperty);
    
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.LastSavedTime.docx", saveOptions);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OoxmlSaveOptions.LastSavedTime.docx");
    System::DateTime lastSavedTimeNew = doc->get_BuiltInDocumentProperties()->get_LastSavedTime();
    
    if (updateLastSavedTimeProperty)
    {
        ASSERT_TRUE((System::DateTime::get_Now() - lastSavedTimeNew).get_Days() < 1);
    }
    else
    {
        ASSERT_EQ(System::DateTime(2021, 5, 11, 6, 32, 0), lastSavedTimeNew);
    }
    //ExEnd
}

namespace gtest_test
{

using ExOoxmlSaveOptions_LastSavedTime_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExOoxmlSaveOptions::LastSavedTime)>::type;

struct ExOoxmlSaveOptions_LastSavedTime : public ExOoxmlSaveOptions, public Aspose::Words::ApiExamples::ExOoxmlSaveOptions, public ::testing::WithParamInterface<ExOoxmlSaveOptions_LastSavedTime_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_LastSavedTime, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LastSavedTime(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_LastSavedTime, ::testing::ValuesIn(ExOoxmlSaveOptions_LastSavedTime::TestCases()));

} // namespace gtest_test

void ExOoxmlSaveOptions::KeepLegacyControlChars(bool keepLegacyControlChars)
{
    //ExStart
    //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
    //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
    //ExSummary:Shows how to support legacy control characters when converting to .docx.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Legacy control character.doc");
    
    // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
    // and then pass it to the document's saving method to modify how we save the document.
    // Set the "KeepLegacyControlChars" property to "true" to preserve
    // the "ShortDateTime" legacy character while saving.
    // Set the "KeepLegacyControlChars" property to "false" to remove
    // the "ShortDateTime" legacy character from the output document.
    auto so = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>(Aspose::Words::SaveFormat::Docx);
    so->set_KeepLegacyControlChars(keepLegacyControlChars);
    
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.KeepLegacyControlChars.docx", so);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OoxmlSaveOptions.KeepLegacyControlChars.docx");
    
    ASSERT_EQ(keepLegacyControlChars ? System::String(u"\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f") : System::String(u"\u001e\f"), doc->get_FirstSection()->get_Body()->GetText());
    //ExEnd
}

namespace gtest_test
{

using ExOoxmlSaveOptions_KeepLegacyControlChars_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExOoxmlSaveOptions::KeepLegacyControlChars)>::type;

struct ExOoxmlSaveOptions_KeepLegacyControlChars : public ExOoxmlSaveOptions, public Aspose::Words::ApiExamples::ExOoxmlSaveOptions, public ::testing::WithParamInterface<ExOoxmlSaveOptions_KeepLegacyControlChars_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_KeepLegacyControlChars, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->KeepLegacyControlChars(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_KeepLegacyControlChars, ::testing::ValuesIn(ExOoxmlSaveOptions_KeepLegacyControlChars::TestCases()));

} // namespace gtest_test

void ExOoxmlSaveOptions::DocumentCompression(Aspose::Words::Saving::CompressionLevel compressionLevel)
{
    //ExStart
    //ExFor:OoxmlSaveOptions.CompressionLevel
    //ExFor:CompressionLevel
    //ExSummary:Shows how to specify the compression level to use while saving an OOXML document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
    // and then pass it to the document's saving method to modify how we save the document.
    // Set the "CompressionLevel" property to "CompressionLevel.Maximum" to apply the strongest and slowest compression.
    // Set the "CompressionLevel" property to "CompressionLevel.Normal" to apply
    // the default compression that Aspose.Words uses while saving OOXML documents.
    // Set the "CompressionLevel" property to "CompressionLevel.Fast" to apply a faster and weaker compression.
    // Set the "CompressionLevel" property to "CompressionLevel.SuperFast" to apply
    // the default compression that Microsoft Word uses.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>(Aspose::Words::SaveFormat::Docx);
    saveOptions->set_CompressionLevel(compressionLevel);
    
    System::SharedPtr<System::Diagnostics::Stopwatch> st = System::Diagnostics::Stopwatch::StartNew();
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.DocumentCompression.docx", saveOptions);
    st->Stop();
    
    auto fileInfo = System::MakeObject<System::IO::FileInfo>(get_ArtifactsDir() + u"OoxmlSaveOptions.DocumentCompression.docx");
    
    std::cout << System::String::Format(u"Saving operation done using the \"{0}\" compression level:", compressionLevel) << std::endl;
    std::cout << System::String::Format(u"\tDuration:\t{0} ms", st->get_ElapsedMilliseconds()) << std::endl;
    std::cout << System::String::Format(u"\tFile Size:\t{0} bytes", fileInfo->get_Length()) << std::endl;
    //ExEnd
    
    int64_t testedFileLength = fileInfo->get_Length();
    
    switch (compressionLevel)
    {
        case Aspose::Words::Saving::CompressionLevel::Maximum:
            ASSERT_TRUE(testedFileLength < 1269000);
            break;
        
        case Aspose::Words::Saving::CompressionLevel::Normal:
            ASSERT_TRUE(testedFileLength < 1271000);
            break;
        
        case Aspose::Words::Saving::CompressionLevel::Fast:
            ASSERT_TRUE(testedFileLength < 1280000);
            break;
        
        case Aspose::Words::Saving::CompressionLevel::SuperFast:
            ASSERT_TRUE(testedFileLength < 1276000);
            break;
        
    }
}

namespace gtest_test
{

using ExOoxmlSaveOptions_DocumentCompression_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExOoxmlSaveOptions::DocumentCompression)>::type;

struct ExOoxmlSaveOptions_DocumentCompression : public ExOoxmlSaveOptions, public Aspose::Words::ApiExamples::ExOoxmlSaveOptions, public ::testing::WithParamInterface<ExOoxmlSaveOptions_DocumentCompression_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Saving::CompressionLevel::Maximum),
            std::make_tuple(Aspose::Words::Saving::CompressionLevel::Fast),
            std::make_tuple(Aspose::Words::Saving::CompressionLevel::Normal),
            std::make_tuple(Aspose::Words::Saving::CompressionLevel::SuperFast),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_DocumentCompression, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->DocumentCompression(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_DocumentCompression, ::testing::ValuesIn(ExOoxmlSaveOptions_DocumentCompression::TestCases()));

} // namespace gtest_test

void ExOoxmlSaveOptions::CheckFileSignatures()
{
    System::ArrayPtr<Aspose::Words::Saving::CompressionLevel> compressionLevels = System::MakeArray<Aspose::Words::Saving::CompressionLevel>({Aspose::Words::Saving::CompressionLevel::Maximum, Aspose::Words::Saving::CompressionLevel::Normal, Aspose::Words::Saving::CompressionLevel::Fast, Aspose::Words::Saving::CompressionLevel::SuperFast});
    
    System::ArrayPtr<System::String> fileSignatures = System::MakeArray<System::String>({u"50 4B 03 04 14 00 02 00 08 00 ", u"50 4B 03 04 14 00 00 00 08 00 ", u"50 4B 03 04 14 00 04 00 08 00 ", u"50 4B 03 04 14 00 06 00 08 00 "});
    
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>(Aspose::Words::SaveFormat::Docx);
    
    int64_t prevFileSize = 0;
    for (int32_t i = 0; i < fileSignatures->get_Length(); i++)
    {
        saveOptions->set_CompressionLevel(compressionLevels[i]);
        doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.CheckFileSignatures.docx", saveOptions);
        
        {
            auto stream = System::MakeObject<System::IO::MemoryStream>();
            {
                System::SharedPtr<System::IO::FileStream> outputFileStream = System::IO::File::Open(get_ArtifactsDir() + u"OoxmlSaveOptions.CheckFileSignatures.docx", System::IO::FileMode::Open);
                int64_t fileSize = outputFileStream->get_Length();
                ASSERT_TRUE(prevFileSize < fileSize);
                
                Aspose::Words::ApiExamples::TestUtil::CopyStream(outputFileStream, stream);
                ASSERT_EQ(fileSignatures[i], Aspose::Words::ApiExamples::TestUtil::DumpArray(stream->ToArray(), 0, 10));
                
                prevFileSize = fileSize;
            }
        }
    }
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, CheckFileSignatures)
{
    s_instance->CheckFileSignatures();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::ExportGeneratorName()
{
    //ExStart
    //ExFor:SaveOptions.ExportGeneratorName
    //ExSummary:Shows how to disable adding name and version of Aspose.Words into produced files.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Use https://docs.aspose.com/words/net/generator-or-producer-name-included-in-output-documents/ to know how to check the result.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_ExportGeneratorName(false);
    
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.ExportGeneratorName.docx", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, ExportGeneratorName)
{
    s_instance->ExportGeneratorName();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::ProgressCallback(Aspose::Words::SaveFormat saveFormat, System::String ext)
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Big document.docx");
    
    // Following formats are supported: Docx, FlatOpc, Docm, Dotm, Dotx.
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>(saveFormat);
    saveOptions->set_ProgressCallback(System::MakeObject<Aspose::Words::ApiExamples::ExOoxmlSaveOptions::SavingProgressCallback>());
    
    System::OperationCanceledException exception; ASPOSE_ASSERT_THROW(static_cast<std::function<void()>>([&doc, &ext, &saveOptions]() -> void
    {
        doc->Save(get_ArtifactsDir() + System::String::Format(u"OoxmlSaveOptions.ProgressCallback.{0}", ext), saveOptions);
    })(), System::OperationCanceledException, &exception);
    System::Nullable<bool> actual = System::Default<System::Nullable<bool>>();
    System::OperationCanceledException condExpression = exception;
    if (condExpression != nullptr)
    {
        actual = condExpression->get_Message().Contains(u"EstimatedProgress");
    }
    ASSERT_TRUE(actual);
}

namespace gtest_test
{

using ExOoxmlSaveOptions_ProgressCallback_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExOoxmlSaveOptions::ProgressCallback)>::type;

struct ExOoxmlSaveOptions_ProgressCallback : public ExOoxmlSaveOptions, public Aspose::Words::ApiExamples::ExOoxmlSaveOptions, public ::testing::WithParamInterface<ExOoxmlSaveOptions_ProgressCallback_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::SaveFormat::Docx, u"docx"),
            std::make_tuple(Aspose::Words::SaveFormat::Docm, u"docm"),
            std::make_tuple(Aspose::Words::SaveFormat::Dotm, u"dotm"),
            std::make_tuple(Aspose::Words::SaveFormat::Dotx, u"dotx"),
            std::make_tuple(Aspose::Words::SaveFormat::FlatOpc, u"flatopc"),
        };
    }
};

TEST_P(ExOoxmlSaveOptions_ProgressCallback, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ProgressCallback(std::get<0>(params), std::get<1>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExOoxmlSaveOptions_ProgressCallback, ::testing::ValuesIn(ExOoxmlSaveOptions_ProgressCallback::TestCases()));

} // namespace gtest_test

void ExOoxmlSaveOptions::Zip64ModeOption()
{
    //ExStart:Zip64ModeOption
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:OoxmlSaveOptions.Zip64Mode
    //ExFor:Zip64Mode
    //ExSummary:Shows how to use ZIP64 format extensions.
    System::Random random;
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    for (int32_t i = 0; i < 10000; i++)
    {
        {
            auto bmp = System::MakeObject<System::Drawing::Bitmap>(5, 5);
            {
                System::SharedPtr<System::Drawing::Graphics> g = System::Drawing::Graphics::FromImage(bmp);
                g->Clear(System::Drawing::Color::FromArgb(random.Next(0, 254), random.Next(0, 254), random.Next(0, 254)));
                {
                    auto ms = System::MakeObject<System::IO::MemoryStream>();
                    bmp->Save(ms, System::Drawing::Imaging::ImageFormat::get_Png());
                    builder->InsertImage(ms->ToArray());
                }
            }
        }
    }
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Zip64Mode(Aspose::Words::Saving::Zip64Mode::Always);
    
    builder->get_Document()->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.Zip64ModeOption.docx", saveOptions);
    //ExEnd:Zip64ModeOption
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, Zip64ModeOption)
{
    s_instance->Zip64ModeOption();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::DigitalSignature()
{
    //ExStart:DigitalSignature
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:OoxmlSaveOptions.DigitalSignatureDetails
    //ExFor:DigitalSignatureDetails
    //ExFor:DigitalSignatureDetails.#ctor(CertificateHolder, SignOptions)
    //ExFor:DigitalSignatureDetails.CertificateHolder
    //ExFor:DigitalSignatureDetails.SignOptions
    //ExSummary:Shows how to sign OOXML document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    System::SharedPtr<Aspose::Words::DigitalSignatures::CertificateHolder> certificateHolder = Aspose::Words::DigitalSignatures::CertificateHolder::Create(get_MyDir() + u"morzal.pfx", u"aw");
    auto signOptions = System::MakeObject<Aspose::Words::DigitalSignatures::SignOptions>();
    signOptions->set_Comments(u"Some comments");
    signOptions->set_SignTime(System::DateTime::get_Now());
    auto digitalSignatureDetails = System::MakeObject<Aspose::Words::Saving::DigitalSignatureDetails>(certificateHolder, signOptions);
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_DigitalSignatureDetails(digitalSignatureDetails);
    
    ASPOSE_ASSERT_EQ(certificateHolder, digitalSignatureDetails->get_CertificateHolder());
    ASSERT_EQ(u"Some comments", digitalSignatureDetails->get_SignOptions()->get_Comments());
    
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.DigitalSignature.docx", saveOptions);
    //ExEnd:DigitalSignature
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, DigitalSignature)
{
    s_instance->DigitalSignature();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::UpdateAmbiguousTextFont()
{
    //ExStart:UpdateAmbiguousTextFont
    //GistId:1a265b92fa0019b26277ecfef3c20330
    //ExFor:SaveOptions.UpdateAmbiguousTextFont
    //ExSummary:Shows how to update the font to match the character code being used.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Special symbol.docx");
    System::SharedPtr<Aspose::Words::Run> run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    std::cout << run->get_Text() << std::endl;
    // ฿
    std::cout << run->get_Font()->get_Name() << std::endl;
    // Arial
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_UpdateAmbiguousTextFont(true);
    doc->Save(get_ArtifactsDir() + u"OoxmlSaveOptions.UpdateAmbiguousTextFont.docx", saveOptions);
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"OoxmlSaveOptions.UpdateAmbiguousTextFont.docx");
    run = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0);
    std::cout << run->get_Text() << std::endl;
    // ฿
    std::cout << run->get_Font()->get_Name() << std::endl;
    // Angsana New
    //ExEnd:UpdateAmbiguousTextFont
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, UpdateAmbiguousTextFont)
{
    s_instance->UpdateAmbiguousTextFont();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
