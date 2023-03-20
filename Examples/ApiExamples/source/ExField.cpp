// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExField.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
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

TEST_F(ExField, DisplayResult)
{
    s_instance->DisplayResult();
}

TEST_F(ExField, CreateWithFieldBuilder)
{
    s_instance->CreateWithFieldBuilder();
}

TEST_F(ExField, RevNum)
{
    s_instance->RevNum();
}

TEST_F(ExField, InsertFieldNone)
{
    s_instance->InsertFieldNone();
}

TEST_F(ExField, InsertTcField)
{
    s_instance->InsertTcField();
}

TEST_F(ExField, InsertTcFieldsAtText)
{
    s_instance->InsertTcFieldsAtText();
}

TEST_F(ExField, FieldLocale)
{
    s_instance->FieldLocale();
}

TEST_F(ExField, InsertFieldWithFieldBuilderException)
{
    s_instance->InsertFieldWithFieldBuilderException();
}

using ExField_PreserveIncludePicture_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExField::PreserveIncludePicture)>::type;

struct ExField_PreserveIncludePicture : public ExField, public ApiExamples::ExField, public ::testing::WithParamInterface<ExField_PreserveIncludePicture_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExField_PreserveIncludePicture, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PreserveIncludePicture(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_PreserveIncludePicture, ::testing::ValuesIn(ExField_PreserveIncludePicture::TestCases()));

TEST_F(ExField, FieldFormat)
{
    s_instance->FieldFormat_();
}

TEST_F(ExField, Unlink)
{
    s_instance->Unlink();
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

TEST_F(ExField, RemoveFields)
{
    s_instance->RemoveFields();
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

TEST_F(ExField, FieldBarcode)
{
    s_instance->FieldBarcode_();
}

TEST_F(ExField, FieldDisplayBarcode)
{
    s_instance->FieldDisplayBarcode_();
}

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

TEST_F(ExField, FieldDate)
{
    s_instance->FieldDate_();
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

TEST_F(ExField, FieldSetRef)
{
    s_instance->FieldSetRef();
}

TEST_F(ExField, FieldTemplate)
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

TEST_F(ExField, SetFieldIndexFormat)
{
    s_instance->SetFieldIndexFormat();
}

using ExField_ConditionEvaluationExtensionPoint_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExField::ConditionEvaluationExtensionPoint)>::type;

struct ExField_ConditionEvaluationExtensionPoint : public ExField,
                                                   public ApiExamples::ExField,
                                                   public ::testing::WithParamInterface<ExField_ConditionEvaluationExtensionPoint_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", 1, nullptr, u"true argument"),
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", 0, nullptr, u"false argument"),
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", -1, u"Custom Error", u"Custom Error"),
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", -1, nullptr, u"true argument"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", 1, nullptr, u"1"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", 0, nullptr, u"0"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", -1, u"Custom Error", u"Custom Error"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", -1, nullptr, u"1"),
        };
    }
};

TEST_P(ExField_ConditionEvaluationExtensionPoint, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ConditionEvaluationExtensionPoint(std::get<0>(params), std::get<1>(params), std::get<2>(params), std::get<3>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_ConditionEvaluationExtensionPoint, ::testing::ValuesIn(ExField_ConditionEvaluationExtensionPoint::TestCases()));

TEST_F(ExField, ComparisonExpressionEvaluatorNestedFields)
{
    s_instance->ComparisonExpressionEvaluatorNestedFields();
}

TEST_F(ExField, ComparisonExpressionEvaluatorHeaderFooterFields)
{
    s_instance->ComparisonExpressionEvaluatorHeaderFooterFields();
}

TEST_F(ExField, FieldUpdatingCallbackTest)
{
    s_instance->FieldUpdatingCallbackTest();
}

}} // namespace ApiExamples::gtest_test
