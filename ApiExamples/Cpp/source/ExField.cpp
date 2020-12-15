// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported
// CPPDEFECT: Aspose.BarCode is not supported
#include "ExField.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/VisitorAction.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldTime.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldTA.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldIncludeText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldNoteRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldPageRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldEQ.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/MailMerge/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <system/collections/dictionary.h>
#include <system/shared_ptr.h>
#include <system/string.h>
#include <system/text/string_builder.h>

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Replacing;
namespace ApiExamples { namespace gtest_test {

class ExField : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExField> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExField>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExField> ExField::s_instance;

TEST_F(ExField, GetFieldFromDocument)
{
    s_instance->GetFieldFromDocument();
}

TEST_F(ExField, GetFieldCode)
{
    s_instance->GetFieldCode();
}

TEST_F(ExField, FieldDisplayResult)
{
    s_instance->FieldDisplayResult();
}

TEST_F(ExField, CreateWithFieldBuilder)
{
    s_instance->CreateWithFieldBuilder();
}

TEST_F(ExField, CreateRevNumFieldByDocumentBuilder)
{
    s_instance->CreateRevNumFieldByDocumentBuilder();
}

TEST_F(ExField, CreateInfoFieldWithFieldBuilder)
{
    s_instance->CreateInfoFieldWithFieldBuilder();
}

TEST_F(ExField, CreateInfoFieldWithDocumentBuilder)
{
    s_instance->CreateInfoFieldWithDocumentBuilder();
}

TEST_F(ExField, GetFieldFromFieldCollection)
{
    s_instance->GetFieldFromFieldCollection();
}

TEST_F(ExField, InsertFieldNone)
{
    s_instance->InsertFieldNone();
}

TEST_F(ExField, InsertTcField)
{
    s_instance->InsertTcField();
}

TEST_F(ExField, FieldLocale)
{
    s_instance->FieldLocale();
}

TEST_F(ExField, ChangeLocale)
{
    s_instance->ChangeLocale();
}

TEST_F(ExField, RemoveTocFromDocument)
{
    s_instance->RemoveTocFromDocument();
}

TEST_F(ExField, InsertTcFieldsAtText)
{
    s_instance->InsertTcFieldsAtText();
}

using UpdateDirtyFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExField::UpdateDirtyFields)>::type;

struct UpdateDirtyFieldsVP : public ExField, public ApiExamples::ExField, public ::testing::WithParamInterface<UpdateDirtyFields_Args>
{
    static std::vector<UpdateDirtyFields_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(UpdateDirtyFieldsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateDirtyFields(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_ExField, UpdateDirtyFieldsVP, ::testing::ValuesIn(UpdateDirtyFieldsVP::TestCases()));

TEST_F(ExField, InsertFieldWithFieldBuilderException)
{
    s_instance->InsertFieldWithFieldBuilderException();
}

TEST_F(ExField, UpdateFieldIgnoringMergeFormat)
{
    s_instance->UpdateFieldIgnoringMergeFormat();
}

TEST_F(ExField, FieldFormat)
{
    s_instance->FieldFormat_();
}

TEST_F(ExField, UnlinkAllFieldsInDocument)
{
    s_instance->UnlinkAllFieldsInDocument();
}

TEST_F(ExField, UnlinkAllFieldsInRange)
{
    s_instance->UnlinkAllFieldsInRange();
}

TEST_F(ExField, UnlinkSingleField)
{
    s_instance->UnlinkSingleField();
}

TEST_F(ExField, UpdateTocPageNumbers)
{
    s_instance->UpdateTocPageNumbers();
}

TEST_F(ExField, DropDownItemCollection)
{
    s_instance->DropDownItemCollection_();
}

TEST_F(ExField, FieldAdvance)
{
    s_instance->FieldAdvance_();
}

TEST_F(ExField, FieldAddressBlock)
{
    s_instance->FieldAddressBlock_();
}

TEST_F(ExField, FieldCollection)
{
    s_instance->FieldCollection_();
}

TEST_F(ExField, FieldCompare)
{
    s_instance->FieldCompare_();
}

TEST_F(ExField, FieldIf)
{
    s_instance->FieldIf_();
}

TEST_F(ExField, FieldAutoNum)
{
    s_instance->FieldAutoNum_();
}

TEST_F(ExField, FieldAutoNumLgl)
{
    s_instance->FieldAutoNumLgl_();
}

TEST_F(ExField, FieldAutoNumOut)
{
    s_instance->FieldAutoNumOut_();
}

TEST_F(ExField, FieldAutoText)
{
    s_instance->FieldAutoText_();
}

TEST_F(ExField, FieldAutoTextList)
{
    s_instance->FieldAutoTextList_();
}

TEST_F(ExField, FieldListNum)
{
    s_instance->FieldListNum_();
}

TEST_F(ExField, FieldToc)
{
    s_instance->FieldToc_();
}

TEST_F(ExField, FieldTocEntryIdentifier)
{
    s_instance->FieldTocEntryIdentifier();
}

TEST_F(ExField, TocSeqPrefix)
{
    s_instance->TocSeqPrefix();
}

TEST_F(ExField, TocSeqNumbering)
{
    s_instance->TocSeqNumbering();
}

TEST_F(ExField, DISABLED_TocSeqBookmark)
{
    s_instance->TocSeqBookmark();
}

TEST_F(ExField, DISABLED_FieldCitation)
{
    s_instance->FieldCitation_();
}

TEST_F(ExField, FieldData)
{
    s_instance->FieldData_();
}

TEST_F(ExField, FieldInclude)
{
    s_instance->FieldInclude_();
}

TEST_F(ExField, FieldIncludePicture)
{
    s_instance->FieldIncludePicture_();
}

TEST_F(ExField, DISABLED_FieldIncludeText)
{
    s_instance->FieldIncludeText_();
}

TEST_F(ExField, DISABLED_FieldHyperlink)
{
    s_instance->FieldHyperlink_();
}

TEST_F(ExField, DISABLED_FieldIndexFilter)
{
    s_instance->FieldIndexFilter();
}

TEST_F(ExField, DISABLED_FieldIndexFormatting)
{
    s_instance->FieldIndexFormatting();
}

TEST_F(ExField, DISABLED_FieldIndexSequence)
{
    s_instance->FieldIndexSequence();
}

TEST_F(ExField, DISABLED_FieldIndexPageNumberSeparator)
{
    s_instance->FieldIndexPageNumberSeparator();
}

TEST_F(ExField, DISABLED_FieldIndexPageRangeBookmark)
{
    s_instance->FieldIndexPageRangeBookmark();
}

TEST_F(ExField, DISABLED_FieldIndexCrossReferenceSeparator)
{
    s_instance->FieldIndexCrossReferenceSeparator();
}

using FieldIndexSubheading_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExField::FieldIndexSubheading)>::type;

struct FieldIndexSubheadingVP : public ExField, public ApiExamples::ExField, public ::testing::WithParamInterface<FieldIndexSubheading_Args>
{
    static std::vector<FieldIndexSubheading_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(FieldIndexSubheadingVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FieldIndexSubheading(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_ExField, FieldIndexSubheadingVP, ::testing::ValuesIn(FieldIndexSubheadingVP::TestCases()));

using FieldIndexYomi_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExField::FieldIndexYomi)>::type;

struct FieldIndexYomiVP : public ExField, public ApiExamples::ExField, public ::testing::WithParamInterface<FieldIndexYomi_Args>
{
    static std::vector<FieldIndexYomi_Args> TestCases()
    {
        return {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(FieldIndexYomiVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FieldIndexYomi(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_ExField, FieldIndexYomiVP, ::testing::ValuesIn(FieldIndexYomiVP::TestCases()));

TEST_F(ExField, FieldBarcode)
{
    s_instance->FieldBarcode_();
}

TEST_F(ExField, FieldDisplayBarcode)
{
    s_instance->FieldDisplayBarcode_();
}

using FieldLinkedObjectsAsText_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExField::FieldLinkedObjectsAsText)>::type;

struct FieldLinkedObjectsAsTextVP : public ExField, public ApiExamples::ExField, public ::testing::WithParamInterface<FieldLinkedObjectsAsText_Args>
{
    static std::vector<FieldLinkedObjectsAsText_Args> TestCases()
    {
        return {
            std::make_tuple(ApiExamples::ExField::InsertLinkedObjectAs::Text),
            std::make_tuple(ApiExamples::ExField::InsertLinkedObjectAs::Unicode),
            std::make_tuple(ApiExamples::ExField::InsertLinkedObjectAs::Html),
            std::make_tuple(ApiExamples::ExField::InsertLinkedObjectAs::Rtf),
        };
    }
};

TEST_P(FieldLinkedObjectsAsTextVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FieldLinkedObjectsAsText(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_ExField, FieldLinkedObjectsAsTextVP, ::testing::ValuesIn(FieldLinkedObjectsAsTextVP::TestCases()));

using FieldLinkedObjectsAsImage_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExField::FieldLinkedObjectsAsImage)>::type;

struct FieldLinkedObjectsAsImageVP : public ExField, public ApiExamples::ExField, public ::testing::WithParamInterface<FieldLinkedObjectsAsImage_Args>
{
    static std::vector<FieldLinkedObjectsAsImage_Args> TestCases()
    {
        return {
            std::make_tuple(ApiExamples::ExField::InsertLinkedObjectAs::Picture),
            std::make_tuple(ApiExamples::ExField::InsertLinkedObjectAs::Bitmap),
        };
    }
};

TEST_P(FieldLinkedObjectsAsImageVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FieldLinkedObjectsAsImage(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_ExField, FieldLinkedObjectsAsImageVP, ::testing::ValuesIn(FieldLinkedObjectsAsImageVP::TestCases()));

TEST_F(ExField, FieldUserAddress)
{
    s_instance->FieldUserAddress_();
}

TEST_F(ExField, FieldUserInitials)
{
    s_instance->FieldUserInitials_();
}

TEST_F(ExField, FieldUserName)
{
    s_instance->FieldUserName_();
}

TEST_F(ExField, DISABLED_FieldStyleRefParagraphNumbers)
{
    s_instance->FieldStyleRefParagraphNumbers();
}

TEST_F(ExField, FieldDate)
{
    s_instance->FieldDate_();
}

TEST_F(ExField, DISABLED_FieldCreateDate)
{
    s_instance->FieldCreateDate_();
}

TEST_F(ExField, DISABLED_FieldSaveDate)
{
    s_instance->FieldSaveDate_();
}

TEST_F(ExField, FieldBuilder)
{
    s_instance->FieldBuilder_();
}

TEST_F(ExField, FieldAuthor)
{
    s_instance->FieldAuthor_();
}

TEST_F(ExField, FieldDocVariable)
{
    s_instance->FieldDocVariable_();
}

TEST_F(ExField, FieldSubject)
{
    s_instance->FieldSubject_();
}

TEST_F(ExField, FieldComments)
{
    s_instance->FieldComments_();
}

TEST_F(ExField, FieldFileSize)
{
    s_instance->FieldFileSize_();
}

TEST_F(ExField, FieldGoToButton)
{
    s_instance->FieldGoToButton_();
}

TEST_F(ExField, FieldFillIn)
{
    s_instance->FieldFillIn_();
}

TEST_F(ExField, FieldInfo)
{
    s_instance->FieldInfo_();
}

TEST_F(ExField, FieldMacroButton)
{
    s_instance->FieldMacroButton_();
}

TEST_F(ExField, FieldKeywords)
{
    s_instance->FieldKeywords_();
}

TEST_F(ExField, FieldNum)
{
    s_instance->FieldNum();
}

TEST_F(ExField, FieldPrint)
{
    s_instance->FieldPrint_();
}

TEST_F(ExField, FieldPrintDate)
{
    s_instance->FieldPrintDate_();
}

TEST_F(ExField, FieldQuote)
{
    s_instance->FieldQuote_();
}

TEST_F(ExField, DISABLED_FieldNoteRef)
{
    s_instance->FieldNoteRef_();
}

TEST_F(ExField, DISABLED_FootnoteRef)
{
    s_instance->FootnoteRef();
}

TEST_F(ExField, DISABLED_FieldPageRef)
{
    s_instance->FieldPageRef_();
}

TEST_F(ExField, DISABLED_FieldRef)
{
    s_instance->FieldRef_();
}

TEST_F(ExField, DISABLED_FieldRD)
{
    s_instance->FieldRD_();
}

TEST_F(ExField, FieldSet)
{
    s_instance->FieldSet_();
}

TEST_F(ExField, DISABLED_FieldTemplate)
{
    s_instance->FieldTemplate_();
}

TEST_F(ExField, FieldSymbol)
{
    s_instance->FieldSymbol_();
}

TEST_F(ExField, FieldTitle)
{
    s_instance->FieldTitle_();
}

TEST_F(ExField, FieldTOA)
{
    s_instance->FieldTOA_();
}

TEST_F(ExField, FieldAddIn)
{
    s_instance->FieldAddIn_();
}

TEST_F(ExField, FieldEditTime)
{
    s_instance->FieldEditTime_();
}

TEST_F(ExField, FieldEQ)
{
    s_instance->FieldEQ_();
}

TEST_F(ExField, FieldForms)
{
    s_instance->FieldForms();
}

TEST_F(ExField, FieldFormula)
{
    s_instance->FieldFormula_();
}

TEST_F(ExField, FieldLastSavedBy)
{
    s_instance->FieldLastSavedBy_();
}

TEST_F(ExField, FieldOcx)
{
    s_instance->FieldOcx_();
}

TEST_F(ExField, FieldPrivate)
{
    s_instance->FieldPrivate_();
}

TEST_F(ExField, FieldSection)
{
    s_instance->FieldSection_();
}

TEST_F(ExField, FieldTime)
{
    s_instance->FieldTime_();
}

TEST_F(ExField, BidiOutline)
{
    s_instance->BidiOutline();
}

TEST_F(ExField, Legacy)
{
    s_instance->Legacy();
}

}} // namespace ApiExamples::gtest_test
