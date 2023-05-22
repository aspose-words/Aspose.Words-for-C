// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocument.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Comparing;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Vba;
using namespace Aspose::Words::WebExtensions;
namespace ApiExamples { namespace gtest_test {

class ExDocument : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExDocument> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExDocument>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExDocument> ExDocument::s_instance;

TEST_F(ExDocument, Constructor)
{
    s_instance->Constructor();
}

TEST_F(ExDocument, LoadFromStream)
{
    s_instance->LoadFromStream();
}

TEST_F(ExDocument, ConvertToPdf)
{
    s_instance->ConvertToPdf();
}

TEST_F(ExDocument, SaveToImageStream)
{
    s_instance->SaveToImageStream();
}

TEST_F(ExDocument, OpenFromStreamWithBaseUri)
{
    s_instance->OpenFromStreamWithBaseUri();
}

TEST_F(ExDocument, LoadEncrypted)
{
    s_instance->LoadEncrypted();
}

TEST_F(ExDocument, TempFolder)
{
    s_instance->TempFolder();
}

TEST_F(ExDocument, ConvertToHtml)
{
    s_instance->ConvertToHtml();
}

TEST_F(ExDocument, ConvertToMhtml)
{
    s_instance->ConvertToMhtml();
}

TEST_F(ExDocument, ConvertToTxt)
{
    s_instance->ConvertToTxt();
}

TEST_F(ExDocument, ConvertToEpub)
{
    s_instance->ConvertToEpub();
}

TEST_F(ExDocument, SaveToStream)
{
    s_instance->SaveToStream();
}

TEST_F(ExDocument, FontChangeViaCallback)
{
    s_instance->FontChangeViaCallback();
}

TEST_F(ExDocument, AppendDocument)
{
    s_instance->AppendDocument();
}

TEST_F(ExDocument, AppendDocumentFromAutomation)
{
    s_instance->AppendDocumentFromAutomation();
}

using ExDocument_ImportList_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::ImportList)>::type;

struct ExDocument_ImportList : public ExDocument, public ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_ImportList_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocument_ImportList, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ImportList(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_ImportList, ::testing::ValuesIn(ExDocument_ImportList::TestCases()));

TEST_F(ExDocument, KeepSourceNumberingSameListIds)
{
    s_instance->KeepSourceNumberingSameListIds();
}

TEST_F(ExDocument, MergePastedLists)
{
    s_instance->MergePastedLists();
}

TEST_F(ExDocument, ValidateIndividualDocumentSignatures)
{
    s_instance->ValidateIndividualDocumentSignatures();
}

TEST_F(ExDocument, DigitalSignature)
{
    s_instance->DigitalSignature_();
}

TEST_F(ExDocument, AppendAllDocumentsInFolder)
{
    s_instance->AppendAllDocumentsInFolder();
}

TEST_F(ExDocument, JoinRunsWithSameFormatting)
{
    s_instance->JoinRunsWithSameFormatting();
}

TEST_F(ExDocument, DefaultTabStop)
{
    s_instance->DefaultTabStop();
}

TEST_F(ExDocument, CloneDocument)
{
    s_instance->CloneDocument();
}

TEST_F(ExDocument, DocumentGetTextToString)
{
    s_instance->DocumentGetTextToString();
}

TEST_F(ExDocument, DocumentByteArray)
{
    s_instance->DocumentByteArray();
}

TEST_F(ExDocument, ProtectUnprotect)
{
    s_instance->ProtectUnprotect();
}

TEST_F(ExDocument, DocumentEnsureMinimum)
{
    s_instance->DocumentEnsureMinimum();
}

TEST_F(ExDocument, RemoveMacrosFromDocument)
{
    s_instance->RemoveMacrosFromDocument();
}

TEST_F(ExDocument, GetPageCount)
{
    s_instance->GetPageCount();
}

TEST_F(ExDocument, GetUpdatedPageProperties)
{
    s_instance->GetUpdatedPageProperties();
}

TEST_F(ExDocument, TableStyleToDirectFormatting)
{
    s_instance->TableStyleToDirectFormatting();
}

TEST_F(ExDocument, GetOriginalFileInfo)
{
    s_instance->GetOriginalFileInfo();
}

TEST_F(ExDocument, FootnoteColumns)
{
    s_instance->FootnoteColumns();
}

TEST_F(ExDocument, Compare)
{
    s_instance->Compare();
}

TEST_F(ExDocument, CompareDocumentWithRevisions)
{
    s_instance->CompareDocumentWithRevisions();
}

TEST_F(ExDocument, CompareOptions)
{
    s_instance->CompareOptions_();
}

using ExDocument_IgnoreDmlUniqueId_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::IgnoreDmlUniqueId)>::type;

struct ExDocument_IgnoreDmlUniqueId : public ExDocument, public ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_IgnoreDmlUniqueId_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_IgnoreDmlUniqueId, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IgnoreDmlUniqueId(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_IgnoreDmlUniqueId, ::testing::ValuesIn(ExDocument_IgnoreDmlUniqueId::TestCases()));

TEST_F(ExDocument, RemoveExternalSchemaReferences)
{
    s_instance->RemoveExternalSchemaReferences();
}

TEST_F(ExDocument, TrackRevisions)
{
    s_instance->TrackRevisions();
}

TEST_F(ExDocument, AcceptAllRevisions)
{
    s_instance->AcceptAllRevisions();
}

TEST_F(ExDocument, GetRevisedPropertiesOfList)
{
    s_instance->GetRevisedPropertiesOfList();
}

TEST_F(ExDocument, UpdateThumbnail)
{
    s_instance->UpdateThumbnail();
}

TEST_F(ExDocument, HyphenationOptions)
{
    s_instance->HyphenationOptions_();
}

TEST_F(ExDocument, HyphenationOptionsDefaultValues)
{
    s_instance->HyphenationOptionsDefaultValues();
}

TEST_F(ExDocument, HyphenationOptionsExceptions)
{
    s_instance->HyphenationOptionsExceptions();
}

TEST_F(ExDocument, OoxmlComplianceVersion)
{
    s_instance->OoxmlComplianceVersion();
}

TEST_F(ExDocument, Cleanup)
{
    s_instance->Cleanup();
}

TEST_F(ExDocument, AutomaticallyUpdateStyles)
{
    s_instance->AutomaticallyUpdateStyles();
}

TEST_F(ExDocument, DefaultTemplate)
{
    s_instance->DefaultTemplate();
}

TEST_F(ExDocument, UseSubstitutions)
{
    s_instance->UseSubstitutions();
}

TEST_F(ExDocument, SetInvalidateFieldTypes)
{
    s_instance->SetInvalidateFieldTypes();
}

TEST_F(ExDocument, LayoutOptionsRevisions)
{
    s_instance->LayoutOptionsRevisions();
}

using ExDocument_LayoutOptionsHiddenText_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::LayoutOptionsHiddenText)>::type;

struct ExDocument_LayoutOptionsHiddenText : public ExDocument,
                                            public ApiExamples::ExDocument,
                                            public ::testing::WithParamInterface<ExDocument_LayoutOptionsHiddenText_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_LayoutOptionsHiddenText, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LayoutOptionsHiddenText(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_LayoutOptionsHiddenText, ::testing::ValuesIn(ExDocument_LayoutOptionsHiddenText::TestCases()));

using ExDocument_LayoutOptionsParagraphMarks_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::LayoutOptionsParagraphMarks)>::type;

struct ExDocument_LayoutOptionsParagraphMarks : public ExDocument,
                                                public ApiExamples::ExDocument,
                                                public ::testing::WithParamInterface<ExDocument_LayoutOptionsParagraphMarks_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_LayoutOptionsParagraphMarks, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->LayoutOptionsParagraphMarks(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_LayoutOptionsParagraphMarks, ::testing::ValuesIn(ExDocument_LayoutOptionsParagraphMarks::TestCases()));

TEST_F(ExDocument, UpdatePageLayout)
{
    s_instance->UpdatePageLayout();
}

TEST_F(ExDocument, DocPackageCustomParts)
{
    s_instance->DocPackageCustomParts();
}

using ExDocument_ShadeFormData_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::ShadeFormData)>::type;

struct ExDocument_ShadeFormData : public ExDocument, public ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_ShadeFormData_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_ShadeFormData, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ShadeFormData(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_ShadeFormData, ::testing::ValuesIn(ExDocument_ShadeFormData::TestCases()));

TEST_F(ExDocument, VersionsCount)
{
    s_instance->VersionsCount();
}

TEST_F(ExDocument, WriteProtection)
{
    s_instance->WriteProtection_();
}

using ExDocument_RemovePersonalInformation_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::RemovePersonalInformation)>::type;

struct ExDocument_RemovePersonalInformation : public ExDocument,
                                              public ApiExamples::ExDocument,
                                              public ::testing::WithParamInterface<ExDocument_RemovePersonalInformation_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_RemovePersonalInformation, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RemovePersonalInformation(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_RemovePersonalInformation, ::testing::ValuesIn(ExDocument_RemovePersonalInformation::TestCases()));

TEST_F(ExDocument, ShowComments)
{
    s_instance->ShowComments();
}

TEST_F(ExDocument, CopyTemplateStylesViaDocument)
{
    s_instance->CopyTemplateStylesViaDocument();
}

TEST_F(ExDocument, CopyTemplateStylesViaDocumentNew)
{
    s_instance->CopyTemplateStylesViaDocumentNew();
}

TEST_F(ExDocument, ReadMacrosFromExistingDocument)
{
    s_instance->ReadMacrosFromExistingDocument();
}

TEST_F(ExDocument, SaveOutputParameters)
{
    s_instance->SaveOutputParameters_();
}

TEST_F(ExDocument, SubDocument)
{
    s_instance->SubDocument_();
}

TEST_F(ExDocument, CreateWebExtension)
{
    s_instance->CreateWebExtension();
}

TEST_F(ExDocument, GetWebExtensionInfo)
{
    s_instance->GetWebExtensionInfo();
}

TEST_F(ExDocument, EpubCover)
{
    s_instance->EpubCover();
}

TEST_F(ExDocument, TextWatermark)
{
    s_instance->TextWatermark();
}

TEST_F(ExDocument, ImageWatermark)
{
    s_instance->ImageWatermark();
}

using ExDocument_SpellingAndGrammarErrors_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::SpellingAndGrammarErrors)>::type;

struct ExDocument_SpellingAndGrammarErrors : public ExDocument,
                                             public ApiExamples::ExDocument,
                                             public ::testing::WithParamInterface<ExDocument_SpellingAndGrammarErrors_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExDocument_SpellingAndGrammarErrors, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SpellingAndGrammarErrors(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_SpellingAndGrammarErrors, ::testing::ValuesIn(ExDocument_SpellingAndGrammarErrors::TestCases()));

using ExDocument_GranularityCompareOption_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::GranularityCompareOption)>::type;

struct ExDocument_GranularityCompareOption : public ExDocument,
                                             public ApiExamples::ExDocument,
                                             public ::testing::WithParamInterface<ExDocument_GranularityCompareOption_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(Granularity::CharLevel),
            std::make_tuple(Granularity::WordLevel),
        };
    }
};

TEST_P(ExDocument_GranularityCompareOption, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->GranularityCompareOption(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_GranularityCompareOption, ::testing::ValuesIn(ExDocument_GranularityCompareOption::TestCases()));

TEST_F(ExDocument, IgnorePrinterMetrics)
{
    s_instance->IgnorePrinterMetrics();
}

TEST_F(ExDocument, ExtractPages)
{
    s_instance->ExtractPages();
}

using ExDocument_SpellingOrGrammar_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::SpellingOrGrammar)>::type;

struct ExDocument_SpellingOrGrammar : public ExDocument, public ApiExamples::ExDocument, public ::testing::WithParamInterface<ExDocument_SpellingOrGrammar_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExDocument_SpellingOrGrammar, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SpellingOrGrammar(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExDocument_SpellingOrGrammar, ::testing::ValuesIn(ExDocument_SpellingOrGrammar::TestCases()));

TEST_F(ExDocument, AllowEmbeddingPostScriptFonts)
{
    s_instance->AllowEmbeddingPostScriptFonts();
}

TEST_F(ExDocument, Frameset)
{
    s_instance->Frameset();
}

}} // namespace ApiExamples::gtest_test
