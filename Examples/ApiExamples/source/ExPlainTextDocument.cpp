// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExPlainTextDocument.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/io/file_stream.h>
#include <system/io/file_mode.h>
#include <system/details/dispose_guard.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/PlainTextDocument.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Saving;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3972757346u, ::Aspose::Words::ApiExamples::ExPlainTextDocument, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExPlainTextDocument : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExPlainTextDocument> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExPlainTextDocument>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExPlainTextDocument> ExPlainTextDocument::s_instance;

} // namespace gtest_test

void ExPlainTextDocument::Load()
{
    //ExStart
    //ExFor:PlainTextDocument
    //ExFor:PlainTextDocument.#ctor(String)
    //ExFor:PlainTextDocument.Text
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"PlainTextDocument.Load.docx");
    
    auto plaintext = System::MakeObject<Aspose::Words::PlainTextDocument>(get_ArtifactsDir() + u"PlainTextDocument.Load.docx");
    
    ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPlainTextDocument, Load)
{
    s_instance->Load();
}

} // namespace gtest_test

void ExPlainTextDocument::LoadFromStream()
{
    //ExStart
    //ExFor:PlainTextDocument.#ctor(Stream)
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext using stream.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    doc->Save(get_ArtifactsDir() + u"PlainTextDocument.LoadFromStream.docx");
    
    {
        auto stream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"PlainTextDocument.LoadFromStream.docx", System::IO::FileMode::Open);
        auto plaintext = System::MakeObject<Aspose::Words::PlainTextDocument>(stream);
        
        ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPlainTextDocument, LoadFromStream)
{
    s_instance->LoadFromStream();
}

} // namespace gtest_test

void ExPlainTextDocument::LoadEncrypted()
{
    //ExStart
    //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
    //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Password(u"MyPassword");
    
    doc->Save(get_ArtifactsDir() + u"PlainTextDocument.LoadEncrypted.docx", saveOptions);
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_Password(u"MyPassword");
    
    auto plaintext = System::MakeObject<Aspose::Words::PlainTextDocument>(get_ArtifactsDir() + u"PlainTextDocument.LoadEncrypted.docx", loadOptions);
    
    ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPlainTextDocument, LoadEncrypted)
{
    s_instance->LoadEncrypted();
}

} // namespace gtest_test

void ExPlainTextDocument::LoadEncryptedUsingStream()
{
    //ExStart
    //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
    //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext using stream.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>();
    saveOptions->set_Password(u"MyPassword");
    
    doc->Save(get_ArtifactsDir() + u"PlainTextDocument.LoadFromStreamWithOptions.docx", saveOptions);
    
    auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
    loadOptions->set_Password(u"MyPassword");
    
    {
        auto stream = System::MakeObject<System::IO::FileStream>(get_ArtifactsDir() + u"PlainTextDocument.LoadFromStreamWithOptions.docx", System::IO::FileMode::Open);
        auto plaintext = System::MakeObject<Aspose::Words::PlainTextDocument>(stream, loadOptions);
        
        ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPlainTextDocument, LoadEncryptedUsingStream)
{
    s_instance->LoadEncryptedUsingStream();
}

} // namespace gtest_test

void ExPlainTextDocument::BuiltInProperties()
{
    //ExStart
    //ExFor:PlainTextDocument.BuiltInDocumentProperties
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's built-in properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    
    doc->Save(get_ArtifactsDir() + u"PlainTextDocument.BuiltInProperties.docx");
    
    auto plaintext = System::MakeObject<Aspose::Words::PlainTextDocument>(get_ArtifactsDir() + u"PlainTextDocument.BuiltInProperties.docx");
    
    ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
    ASSERT_EQ(u"John Doe", plaintext->get_BuiltInDocumentProperties()->get_Author());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPlainTextDocument, BuiltInProperties)
{
    s_instance->BuiltInProperties();
}

} // namespace gtest_test

void ExPlainTextDocument::CustomDocumentProperties()
{
    //ExStart
    //ExFor:PlainTextDocument.CustomDocumentProperties
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's custom properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln(u"Hello world!");
    doc->get_CustomDocumentProperties()->Add(u"Location of writing", System::String(u"123 Main St, London, UK"));
    
    doc->Save(get_ArtifactsDir() + u"PlainTextDocument.CustomDocumentProperties.docx");
    
    auto plaintext = System::MakeObject<Aspose::Words::PlainTextDocument>(get_ArtifactsDir() + u"PlainTextDocument.CustomDocumentProperties.docx");
    
    ASSERT_EQ(u"Hello world!", plaintext->get_Text().Trim());
    ASPOSE_ASSERT_EQ(u"123 Main St, London, UK", plaintext->get_CustomDocumentProperties()->idx_get(u"Location of writing")->get_Value());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExPlainTextDocument, CustomDocumentProperties)
{
    s_instance->CustomDocumentProperties();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
