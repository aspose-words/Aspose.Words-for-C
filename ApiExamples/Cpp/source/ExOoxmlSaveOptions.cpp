// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExOoxmlSaveOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/date_time.h>
#include <gtest/gtest.h>
#include <functional>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeMarkupLanguage.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace ApiExamples {

namespace gtest_test
{

class ExOoxmlSaveOptions : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExOoxmlSaveOptions> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExOoxmlSaveOptions>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExOoxmlSaveOptions> ExOoxmlSaveOptions::s_instance;

} // namespace gtest_test

void ExOoxmlSaveOptions::Password()
{
    //ExStart
    //ExFor:OoxmlSaveOptions.Password
    //ExSummary:Shows how to create a password protected Office Open XML document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");

    // Create a SaveOptions object with a password and save our document with it
    auto saveOptions = MakeObject<OoxmlSaveOptions>();
    saveOptions->set_Password(u"MyPassword");

    doc->Save(ArtifactsDir + u"OoxmlSaveOptions.Password.docx", saveOptions);

    // This document cannot be opened like a normal document

    auto openDocument = [&doc]()
    {
        doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Password.docx");
    };

    ASSERT_THROW(openDocument(), IncorrectPasswordException);

    // We can open the document and access its contents by passing the correct password to a LoadOptions object
    doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Password.docx", MakeObject<LoadOptions>(u"MyPassword"));

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
    //ExFor:OoxmlSaveOptions
    //ExFor:OoxmlSaveOptions.#ctor()
    //ExFor:OoxmlSaveOptions.SaveFormat
    //ExFor:OoxmlCompliance
    //ExFor:OoxmlSaveOptions.Compliance
    //ExFor:ShapeMarkupLanguage
    //ExSummary:Shows conversion VML shapes to DML using ISO/IEC 29500:2008 Strict compliance level.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set Word2003 version for document, for inserting image as VML shape
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2003);
    builder->InsertImage(ImageDir + u"Transparent background logo.png");

    ASSERT_EQ(Aspose::Words::Drawing::ShapeMarkupLanguage::Vml, (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_MarkupLanguage());

    // Iso29500_2008 does not allow VML shapes
    // You need to use OoxmlCompliance.Iso29500_2008_Strict for converting VML to DML shapes
    auto saveOptions = MakeObject<OoxmlSaveOptions>();
    saveOptions->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Strict);
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Docx);

    doc->Save(ArtifactsDir + u"OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

    // The markup language of our shape has changed according to the compliance type
    doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.Iso29500Strict.docx");

    ASSERT_EQ(Aspose::Words::Drawing::ShapeMarkupLanguage::Dml, (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_MarkupLanguage());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, Iso29500Strict)
{
    s_instance->Iso29500Strict();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::RestartingDocumentList(bool doRestartListAtEachSection)
{
    //ExStart
    //ExFor:List.IsRestartAtEachSection
    //ExSummary:Shows how to specify that the list has to be restarted at each section.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberDefault);

    SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->idx_get(0);

    // Set true to specify that the list has to be restarted at each section
    list->set_IsRestartAtEachSection(doRestartListAtEachSection);

    // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
    auto options = MakeObject<OoxmlSaveOptions>();
    options->set_Compliance(Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Transitional);

    builder->get_ListFormat()->set_List(list);

    builder->Writeln(u"List item 1");
    builder->Writeln(u"List item 2");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Writeln(u"List item 3");
    builder->Writeln(u"List item 4");

    doc->Save(ArtifactsDir + u"OoxmlSaveOptions.RestartingDocumentList.docx", options);
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.RestartingDocumentList.docx");

    ASPOSE_ASSERT_EQ(doRestartListAtEachSection, doc->get_Lists()->idx_get(0)->get_IsRestartAtEachSection());
}

namespace gtest_test
{

using RestartingDocumentList_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::RestartingDocumentList)>::type;

struct RestartingDocumentListVP : public ExOoxmlSaveOptions, public ApiExamples::ExOoxmlSaveOptions, public ::testing::WithParamInterface<RestartingDocumentList_Args>
{
    static std::vector<RestartingDocumentList_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(RestartingDocumentListVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RestartingDocumentList(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExOoxmlSaveOptions, RestartingDocumentListVP, ::testing::ValuesIn(RestartingDocumentListVP::TestCases()));

} // namespace gtest_test

void ExOoxmlSaveOptions::UpdatingLastSavedTimeDocument()
{
    //ExStart
    //ExFor:SaveOptions.UpdateLastSavedTimeProperty
    //ExSummary:Shows how to update a document time property when you want to save it.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // Get last saved time
    System::DateTime documentTimeBeforeSave = doc->get_BuiltInDocumentProperties()->get_LastSavedTime();

    auto saveOptions = MakeObject<OoxmlSaveOptions>();
    saveOptions->set_UpdateLastSavedTimeProperty(true);

    doc->Save(ArtifactsDir + u"OoxmlSaveOptions.UpdatingLastSavedTimeDocument.docx", saveOptions);
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    System::DateTime documentTimeAfterSave = doc->get_BuiltInDocumentProperties()->get_LastSavedTime();

    ASSERT_TRUE(documentTimeBeforeSave < documentTimeAfterSave);
}

namespace gtest_test
{

TEST_F(ExOoxmlSaveOptions, UpdatingLastSavedTimeDocument)
{
    s_instance->UpdatingLastSavedTimeDocument();
}

} // namespace gtest_test

void ExOoxmlSaveOptions::KeepLegacyControlChars(bool doKeepLegacyControlChars)
{
    //ExStart
    //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
    //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
    //ExSummary:Shows how to support legacy control characters when converting to .docx.
    auto doc = MakeObject<Document>(MyDir + u"Legacy control character.doc");

    // Note that only one legacy character (ShortDateTime) is supported which declared in the "DOC" format
    auto so = MakeObject<OoxmlSaveOptions>(Aspose::Words::SaveFormat::Docx);
    so->set_KeepLegacyControlChars(doKeepLegacyControlChars);

    doc->Save(ArtifactsDir + u"OoxmlSaveOptions.KeepLegacyControlChars.docx", so);

    // Open the saved document and verify results
    doc = MakeObject<Document>(ArtifactsDir + u"OoxmlSaveOptions.KeepLegacyControlChars.docx");

    if (doKeepLegacyControlChars)
    {
        ASSERT_EQ(u"\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f", doc->get_FirstSection()->get_Body()->GetText());
    }
    else
    {
        ASSERT_EQ(u"\u001e\f", doc->get_FirstSection()->get_Body()->GetText());
    }
    //ExEnd
}

namespace gtest_test
{

using KeepLegacyControlChars_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExOoxmlSaveOptions::KeepLegacyControlChars)>::type;

struct KeepLegacyControlCharsVP : public ExOoxmlSaveOptions, public ApiExamples::ExOoxmlSaveOptions, public ::testing::WithParamInterface<KeepLegacyControlChars_Args>
{
    static std::vector<KeepLegacyControlChars_Args> TestCases()
    {
        return
        {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(KeepLegacyControlCharsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->KeepLegacyControlChars(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExOoxmlSaveOptions, KeepLegacyControlCharsVP, ::testing::ValuesIn(KeepLegacyControlCharsVP::TestCases()));

} // namespace gtest_test

} // namespace ApiExamples
