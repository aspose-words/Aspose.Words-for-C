#include "ExField.h"

#include <xml/xml_node_list.h>
#include <xml/xml_node.h>
#include <xml/xml_namespace_manager.h>
#include <xml/xml_document.h>
#include <xml/xml_attribute_collection.h>
#include <testing/test_predicates.h>
#include <system/timezone_info.h>
#include <system/timespan.h>
#include <system/threading/thread.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/setters_wrap.h>
#include <system/primitive_types.h>
#include <system/object_ext.h>
#include <system/linq/enumerable.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/globalization/um_al_qura_calendar.h>
#include <system/globalization/date_time_format_info.h>
#include <system/globalization/culture_info.h>
#include <system/globalization/calendar.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/convert.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/action.h>
#include <iostream>
#include <functional>
#include <drawing/color_translator.h>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Tables/TableCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Math/OfficeMath.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Loading/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/UserInformation.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldChar.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/GeneralFormatCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/GeneralFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/Format/FieldFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldUpdateCultureSource.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserName.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserInitials.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserAddress.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldInfo.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldDisplayBarcode.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldBidiOutline.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Transitional/FieldBarcode.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldUnknown.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldOcx.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldData.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Other/FieldAddIn.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldSeq.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldSectionPages.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldSection.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldRevNum.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldPage.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldListNum.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldAutoNumOut.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldAutoNumLgl.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/Numbering/FieldAutoNum.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/MergeFieldImageDimension.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldAddressBlock.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldStyleRef.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldSet.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldQuote.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldLink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldIncludePicture.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldInclude.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldImport.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldGlossary.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldFillIn.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldDdeAuto.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldDde.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldAutoTextList.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldAutoText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/Index/FieldIndexFormat.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldXE.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToc.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldToa.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldTC.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldRD.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/IndexAndTables/FieldIndex.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/FormFields/FieldFormText.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/FormFields/FieldFormDropDown.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/FormFields/FieldFormCheckBox.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldSymbol.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldFormula.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/EquationsAndFormulas/FieldAdvance.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldTitle.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldTemplate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldSubject.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldPrivate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldNumWords.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldNumPages.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldNumChars.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldLastSavedBy.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldKeywords.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldFileSize.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldDocProperty.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldComments.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldAuthor.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldPrint.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldMacroButton.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldIfComparisonResult.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldIf.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldGoToButton.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldDocVariable.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentAutomation/FieldCompare.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldSaveDate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldPrintDate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldEditTime.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldDate.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldCreateDate.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldBuilder/FieldBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldBuilder/FieldArgumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/VariableCollection.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockGallery.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockCollection.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlockBehavior.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/BuildingBlock.h>
#include <Aspose.Words.Cpp/Model/Borders/Shading.h>
#include <Aspose.Words.Cpp/Model/Borders/BorderCollection.h>
#include <Aspose.Words.Cpp/Model/Borders/Border.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkCollection.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/Bookmark.h>
#include <Aspose.Words.Cpp/Model/Bibliography/SourceType.h>
#include <Aspose.Words.Cpp/Model/Bibliography/Source.h>
#include <Aspose.Words.Cpp/Model/Bibliography/PersonCollection.h>
#include <Aspose.Words.Cpp/Model/Bibliography/Person.h>
#include <Aspose.Words.Cpp/Model/Bibliography/ContributorCollection.h>
#include <Aspose.Words.Cpp/Model/Bibliography/Contributor.h>
#include <Aspose.Words.Cpp/Model/Bibliography/Bibliography.h>

#include "DocumentHelper.h"
#include "TestUtil.h"


using namespace Aspose::Words::Bibliography;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;
using namespace Aspose::Words::Math;
using namespace Aspose::Words::Notes;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3824201890u, ::Aspose::Words::ApiExamples::ExField::InsertTcFieldHandler, ThisTypeBaseTypesInfo);

ExField::InsertTcFieldHandler::InsertTcFieldHandler(System::String text, System::String switches)
{
    mFieldText = text;
    mFieldSwitches = switches;
}

Aspose::Words::Replacing::ReplaceAction ExField::InsertTcFieldHandler::Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> args)
{
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(System::ExplicitCast<Aspose::Words::Document>(args->get_MatchNode()->get_Document()));
    builder->MoveTo(args->get_MatchNode());
    
    // If the user-specified text is used in the field as display text, use that, otherwise
    // use the match String as the display text.
    System::String insertText = !System::String::IsNullOrEmpty(mFieldText) ? mFieldText : args->get_Match()->get_Value();
    
    // Insert the TC field before this node using the specified String
    // as the display text and user-defined switches.
    builder->InsertField(System::String::Format(u"TC \"{0}\" {1}", insertText, mFieldSwitches));
    
    return ReplaceAction::Skip;
}

RTTI_INFO_IMPL_HASH(652372560u, ::Aspose::Words::ApiExamples::ExField::SqliteFieldDatabaseProvider, ThisTypeBaseTypesInfo);

RTTI_INFO_IMPL_HASH(2554972784u, ::Aspose::Words::ApiExamples::ExField::MyPromptRespondent, ThisTypeBaseTypesInfo);

System::String ExField::MyPromptRespondent::Respond(System::String promptText, System::String defaultResponse)
{
    return System::String(u"Response from MyPromptRespondent. ") + defaultResponse;
}

RTTI_INFO_IMPL_HASH(1653659838u, ::Aspose::Words::ApiExamples::ExField::FieldVisitor, ThisTypeBaseTypesInfo);

ExField::FieldVisitor::FieldVisitor()
{
    mBuilder = System::MakeObject<System::Text::StringBuilder>();
}

System::String ExField::FieldVisitor::GetText()
{
    return mBuilder->ToString();
}

Aspose::Words::VisitorAction ExField::FieldVisitor::VisitFieldStart(System::SharedPtr<Aspose::Words::Fields::FieldStart> fieldStart)
{
    mBuilder->AppendLine(System::String(u"Found field: ") + System::ObjectExt::ToString(fieldStart->get_FieldType()));
    mBuilder->AppendLine(System::String(u"\tField code: ") + fieldStart->GetField()->GetFieldCode());
    mBuilder->AppendLine(System::String(u"\tDisplayed as: ") + fieldStart->GetField()->get_Result());
    
    return VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExField::FieldVisitor::VisitFieldSeparator(System::SharedPtr<Aspose::Words::Fields::FieldSeparator> fieldSeparator)
{
    mBuilder->AppendLine(System::String(u"\tFound separator: ") + fieldSeparator->GetText());
    
    return VisitorAction::Continue;
}

Aspose::Words::VisitorAction ExField::FieldVisitor::VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd)
{
    mBuilder->AppendLine(System::String(u"End of field: ") + System::ObjectExt::ToString(fieldEnd->get_FieldType()));
    
    return VisitorAction::Continue;
}

RTTI_INFO_IMPL_HASH(3923600127u, ::Aspose::Words::ApiExamples::ExField::BibliographyStylesProvider, ThisTypeBaseTypesInfo);

System::SharedPtr<System::IO::Stream> ExField::BibliographyStylesProvider::GetStyle(System::String styleFileName)
{
    return System::IO::File::OpenRead(get_MyDir() + u"Bibliography custom style.xsl");
}

RTTI_INFO_IMPL_HASH(3072805081u, ::Aspose::Words::ApiExamples::ExField::MergedImageResizer, ThisTypeBaseTypesInfo);

ExField::MergedImageResizer::MergedImageResizer(double imageWidth, double imageHeight, Aspose::Words::Fields::MergeFieldImageDimensionUnit unit)
    : mImageWidth(0), mImageHeight(0), mUnit((Aspose::Words::Fields::MergeFieldImageDimensionUnit)0)
{
    mImageWidth = imageWidth;
    mImageHeight = imageHeight;
    mUnit = unit;
}

void ExField::MergedImageResizer::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> e)
{
    throw System::NotImplementedException();
}

void ExField::MergedImageResizer::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    args->set_ImageFileName(System::ObjectExt::ToString(args->get_FieldValue()));
    args->set_ImageWidth(System::MakeObject<Aspose::Words::Fields::MergeFieldImageDimension>(mImageWidth, mUnit));
    args->set_ImageHeight(System::MakeObject<Aspose::Words::Fields::MergeFieldImageDimension>(mImageHeight, mUnit));
    
    ASPOSE_ASSERT_EQ(args->get_ImageWidth()->get_Value(), mImageWidth);
    ASSERT_EQ(args->get_ImageWidth()->get_Unit(), mUnit);
    ASPOSE_ASSERT_EQ(args->get_ImageHeight()->get_Value(), mImageHeight);
    ASSERT_EQ(args->get_ImageHeight()->get_Unit(), mUnit);
    ASSERT_TRUE(System::TestTools::IsNull(args->get_Shape()));
}

RTTI_INFO_IMPL_HASH(3818366749u, ::Aspose::Words::ApiExamples::ExField::ImageFilenameCallback, ThisTypeBaseTypesInfo);

ExField::ImageFilenameCallback::ImageFilenameCallback()
{
    mImageFilenames = System::MakeObject<System::Collections::Generic::Dictionary<System::String, System::String>>();
    mImageFilenames->Add(u"Dark logo", get_ImageDir() + u"Logo.jpg");
    mImageFilenames->Add(u"Transparent logo", get_ImageDir() + u"Transparent background logo.png");
}

void ExField::ImageFilenameCallback::FieldMerging(System::SharedPtr<Aspose::Words::MailMerging::FieldMergingArgs> args)
{
    throw System::NotImplementedException();
}

void ExField::ImageFilenameCallback::ImageFieldMerging(System::SharedPtr<Aspose::Words::MailMerging::ImageFieldMergingArgs> args)
{
    if (mImageFilenames->ContainsKey(System::ObjectExt::ToString(args->get_FieldValue())))
    {
    }
    
    ASSERT_FALSE(System::TestTools::IsNull(args->get_Image()));
}

RTTI_INFO_IMPL_HASH(3477670916u, ::Aspose::Words::ApiExamples::ExField::PromptRespondent, ThisTypeBaseTypesInfo);

System::String ExField::PromptRespondent::Respond(System::String promptText, System::String defaultResponse)
{
    return System::String(u"Response modified by PromptRespondent. ") + defaultResponse;
}

RTTI_INFO_IMPL_HASH(2634561083u, ::Aspose::Words::ApiExamples::ExField::FieldPrivateRemover, ThisTypeBaseTypesInfo);

ExField::FieldPrivateRemover::FieldPrivateRemover() : mFieldsRemovedCount(0)
{
    mFieldsRemovedCount = 0;
}

int32_t ExField::FieldPrivateRemover::GetFieldsRemovedCount()
{
    return mFieldsRemovedCount;
}

Aspose::Words::VisitorAction ExField::FieldPrivateRemover::VisitFieldEnd(System::SharedPtr<Aspose::Words::Fields::FieldEnd> fieldEnd)
{
    if (fieldEnd->get_FieldType() == FieldType::FieldPrivate)
    {
        fieldEnd->GetField()->Remove();
        mFieldsRemovedCount++;
    }
    
    return VisitorAction::Continue;
}

RTTI_INFO_IMPL_HASH(4103448900u, ::Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator, ThisTypeBaseTypesInfo);

ExField::ComparisonExpressionEvaluator::ComparisonExpressionEvaluator(System::SharedPtr<Aspose::Words::Fields::ComparisonEvaluationResult> result)
    : mInvocations(System::MakeObject<System::Collections::Generic::List<System::ArrayPtr<System::String>>>())
{
    mResult = result;
    if (mResult != nullptr)
    {
        std::cout << mResult->get_ErrorMessage() << std::endl;
        std::cout << System::Convert::ToString(mResult->get_Result()) << std::endl;
    }
}

System::SharedPtr<Aspose::Words::Fields::ComparisonEvaluationResult> ExField::ComparisonExpressionEvaluator::Evaluate(System::SharedPtr<Aspose::Words::Fields::Field> field, System::SharedPtr<Aspose::Words::Fields::ComparisonExpression> expression)
{
    mInvocations->Add(System::MakeArray<System::String>({expression->get_LeftExpression(), 
        expression->get_ComparisonOperator(), expression->get_RightExpression()}));
    
    return mResult;
}

System::SharedPtr<Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator> ExField::ComparisonExpressionEvaluator::AssertInvocationsCount(int32_t expected)
{
    EXPECT_EQ(mInvocations->get_Count(), expected);
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator> ExField::ComparisonExpressionEvaluator::AssertInvocationArguments(int32_t invocationIndex, System::String expectedLeftExpression, System::String expectedComparisonOperator, System::String expectedRightExpression)
{
    System::ArrayPtr<System::String> arguments = mInvocations->idx_get(invocationIndex);
    
    EXPECT_EQ(arguments[0], expectedLeftExpression);
    EXPECT_EQ(arguments[1], expectedComparisonOperator);
    EXPECT_EQ(arguments[2], expectedRightExpression);
    
    return System::MakeSharedPtr(this);
}

ExField::ComparisonExpressionEvaluator::~ComparisonExpressionEvaluator()
{
}

RTTI_INFO_IMPL_HASH(1515001563u, ::Aspose::Words::ApiExamples::ExField::FieldUpdatingCallback, ThisTypeBaseTypesInfo);

System::SharedPtr<System::Collections::Generic::IList<System::String>> ExField::FieldUpdatingCallback::get_FieldUpdatedCalls() const
{
    return pr_FieldUpdatedCalls;
}

ExField::FieldUpdatingCallback::FieldUpdatingCallback()
{
    pr_FieldUpdatedCalls = System::MakeObject<System::Collections::Generic::List<System::String>>();
}

void ExField::FieldUpdatingCallback::FieldUpdating(System::SharedPtr<Aspose::Words::Fields::Field> field)
{
    if (field->get_Type() == FieldType::FieldAuthor)
    {
        auto fieldAuthor = System::ExplicitCast<Aspose::Words::Fields::FieldAuthor>(field);
        fieldAuthor->set_AuthorName(u"Updating John Doe");
    }
}

void ExField::FieldUpdatingCallback::FieldUpdated(System::SharedPtr<Aspose::Words::Fields::Field> field)
{
    get_FieldUpdatedCalls()->Add(field->get_Result());
}

void ExField::FieldUpdatingCallback::Notify(System::SharedPtr<Aspose::Words::Fields::FieldUpdatingProgressArgs> args)
{
    std::cout << System::String::Format(u"{0}/{1}", args->get_UpdateCompleted(), args->get_TotalFieldsCount()) << std::endl;
    std::cout << System::String::Format(u"{0}", args->get_UpdatedFieldsCount()) << std::endl;
}


RTTI_INFO_IMPL_HASH(2713581834u, ::Aspose::Words::ApiExamples::ExField, ThisTypeBaseTypesInfo);

void ExField::RemoveSequence(System::SharedPtr<Aspose::Words::Node> start, System::SharedPtr<Aspose::Words::Node> end)
{
    System::SharedPtr<Aspose::Words::Node> curNode = start->NextPreOrder(start->get_Document());
    while (curNode != nullptr && !System::ObjectExt::Equals(curNode, end))
    {
        System::SharedPtr<Aspose::Words::Node> nextNode = curNode->NextPreOrder(start->get_Document());
        
        if (curNode->get_IsComposite())
        {
            auto curComposite = System::ExplicitCast<Aspose::Words::CompositeNode>(curNode);
            if (!curComposite->GetChildNodes(NodeType::Any, true)->Contains(end) && !curComposite->GetChildNodes(NodeType::Any, true)->Contains(start))
            {
                nextNode = curNode->get_NextSibling();
                curNode->Remove();
            }
        }
        else
        {
            curNode->Remove();
        }
        
        curNode = nextNode;
    }
}

void ExField::TestFieldCollection(System::String fieldVisitorText)
{
    ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldDate"));
    ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldTime"));
    ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldRevisionNum"));
    ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldAuthor"));
    ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldSubject"));
    ASSERT_TRUE(fieldVisitorText.Contains(u"Found field: FieldQuote"));
}

void ExField::InsertNumberedClause(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String heading, System::String contents, Aspose::Words::StyleIdentifier headingStyle)
{
    builder->InsertField(FieldType::FieldAutoNumLegal, true);
    builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleIdentifier(headingStyle);
    builder->Writeln(heading);
    
    // This text will belong to the auto num legal field above it.
    // It will collapse when we click the arrow next to the corresponding AUTONUMLGL field in Microsoft Word.
    builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::BodyText);
    builder->Writeln(contents);
}

void ExField::TestFieldAutoNumLgl(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    
    for (auto&& field : System::IterateOver<Aspose::Words::Fields::FieldAutoNumLgl>(doc->get_Range()->get_Fields()->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
    {
        return f->get_Type() == FieldType::FieldAutoNumLegal;
    })))->LINQ_ToList()))
    {
        TestUtil::VerifyField(FieldType::FieldAutoNumLegal, u" AUTONUMLGL  \\s : \\e", System::String::Empty, field);
        
        ASSERT_EQ(field->get_SeparatorCharacter(), u":");
        ASSERT_TRUE(field->get_RemoveTrailingPeriod());
    }
}

void ExField::AppendAutoTextEntry(System::SharedPtr<Aspose::Words::BuildingBlocks::GlossaryDocument> glossaryDoc, System::String name, System::String contents)
{
    auto buildingBlock = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(glossaryDoc);
    buildingBlock->set_Name(name);
    buildingBlock->set_Gallery(BuildingBlockGallery::AutoText);
    buildingBlock->set_Category(u"General");
    buildingBlock->set_Behavior(BuildingBlockBehavior::Paragraph);
    
    auto section = System::MakeObject<Aspose::Words::Section>(glossaryDoc);
    section->AppendChild<System::SharedPtr<Aspose::Words::Body>>(System::MakeObject<Aspose::Words::Body>(glossaryDoc));
    section->get_Body()->AppendParagraph(contents);
    buildingBlock->AppendChild<System::SharedPtr<Aspose::Words::Section>>(section);
    
    glossaryDoc->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(buildingBlock);
}

void ExField::TestFieldAutoTextList(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(doc->get_GlossaryDocument()->get_Count(), 3);
    ASSERT_EQ(doc->get_GlossaryDocument()->get_BuildingBlocks()->idx_get(0)->get_Name(), u"AutoText 1");
    ASSERT_EQ(doc->get_GlossaryDocument()->get_BuildingBlocks()->idx_get(0)->GetText().Trim(), u"Contents of AutoText 1");
    ASSERT_EQ(doc->get_GlossaryDocument()->get_BuildingBlocks()->idx_get(1)->get_Name(), u"AutoText 2");
    ASSERT_EQ(doc->get_GlossaryDocument()->get_BuildingBlocks()->idx_get(1)->GetText().Trim(), u"Contents of AutoText 2");
    ASSERT_EQ(doc->get_GlossaryDocument()->get_BuildingBlocks()->idx_get(2)->get_Name(), u"AutoText 3");
    ASSERT_EQ(doc->get_GlossaryDocument()->get_BuildingBlocks()->idx_get(2)->GetText().Trim(), u"Contents of AutoText 3");
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAutoTextList>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldAutoTextList, u" AUTOTEXTLIST  \"Right click here to select an AutoText block\" \\s \"Heading 1\" \\t \"Hover tip text for AutoTextList goes here\"", System::String::Empty, field);
    ASSERT_EQ(field->get_EntryName(), u"Right click here to select an AutoText block");
    ASSERT_EQ(field->get_ListStyle(), u"Heading 1");
    ASSERT_EQ(field->get_ScreenTip(), u"Hover tip text for AutoTextList goes here");
}

void ExField::InsertNewPageWithHeading(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String captionText, System::String styleName)
{
    builder->InsertBreak(BreakType::PageBreak);
    System::String originalStyle = builder->get_ParagraphFormat()->get_StyleName();
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(styleName));
    builder->Writeln(captionText);
    builder->get_ParagraphFormat()->set_Style(builder->get_Document()->get_Styles()->idx_get(originalStyle));
}

void ExField::TestFieldToc(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark");
    ASSERT_EQ(field->get_CustomStyles(), u"Quote; 6; Intense Quote; 7");
    ASSERT_EQ(field->get_EntrySeparator(), u"-");
    ASSERT_EQ(field->get_HeadingLevelRange(), u"1-3");
    ASSERT_EQ(field->get_PageNumberOmittingLevelRange(), u"2-5");
    ASSERT_FALSE(field->get_HideInWebLayout());
    ASSERT_TRUE(field->get_InsertHyperlinks());
    ASSERT_TRUE(field->get_PreserveLineBreaks());
    ASSERT_TRUE(field->get_PreserveTabs());
    ASSERT_TRUE(field->UpdatePageNumbers());
    ASSERT_FALSE(field->get_UseParagraphOutlineLevel());
    ASSERT_EQ(field->GetFieldCode(), u" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w");
    ASSERT_EQ(field->get_Result(), System::String(u"\u0013 HYPERLINK \\l \"_Toc256000001\" \u0014First entry-\u0013 PAGEREF _Toc256000001 \\h \u00142\u0015\u0015\r") + u"\u0013 HYPERLINK \\l \"_Toc256000002\" \u0014Second entry-\u0013 PAGEREF _Toc256000002 \\h \u00143\u0015\u0015\r" + u"\u0013 HYPERLINK \\l \"_Toc256000003\" \u0014Third entry-\u0013 PAGEREF _Toc256000003 \\h \u00144\u0015\u0015\r" + u"\u0013 HYPERLINK \\l \"_Toc256000004\" \u0014Fourth entry-\u0013 PAGEREF _Toc256000004 \\h \u00145\u0015\u0015\r" + u"\u0013 HYPERLINK \\l \"_Toc256000005\" \u0014Fifth entry\u0015\r" + u"\u0013 HYPERLINK \\l \"_Toc256000006\" \u0014Sixth entry\u0015\r");
}

void ExField::InsertTocEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String text, System::String typeIdentifier, System::String entryLevel)
{
    auto fieldTc = System::ExplicitCast<Aspose::Words::Fields::FieldTC>(builder->InsertField(FieldType::FieldTOCEntry, true));
    fieldTc->set_OmitPageNumber(true);
    fieldTc->set_Text(text);
    fieldTc->set_TypeIdentifier(typeIdentifier);
    fieldTc->set_EntryLevel(entryLevel);
}

void ExField::TestFieldTocEntryIdentifier(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    auto fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldTOC, u" TOC  \\f A \\l 1-3", u"TC field 1\rTC field 2\r", fieldToc);
    ASSERT_EQ(fieldToc->get_EntryIdentifier(), u"A");
    ASSERT_EQ(fieldToc->get_EntryLevelRange(), u"1-3");
    
    auto fieldTc = System::ExplicitCast<Aspose::Words::Fields::FieldTC>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 1\" \\n \\f A \\l 1", System::String::Empty, fieldTc);
    ASSERT_TRUE(fieldTc->get_OmitPageNumber());
    ASSERT_EQ(fieldTc->get_Text(), u"TC field 1");
    ASSERT_EQ(fieldTc->get_TypeIdentifier(), u"A");
    ASSERT_EQ(fieldTc->get_EntryLevel(), u"1");
    
    fieldTc = System::ExplicitCast<Aspose::Words::Fields::FieldTC>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 2\" \\n \\f A \\l 2", System::String::Empty, fieldTc);
    ASSERT_TRUE(fieldTc->get_OmitPageNumber());
    ASSERT_EQ(fieldTc->get_Text(), u"TC field 2");
    ASSERT_EQ(fieldTc->get_TypeIdentifier(), u"A");
    ASSERT_EQ(fieldTc->get_EntryLevel(), u"2");
    
    fieldTc = System::ExplicitCast<Aspose::Words::Fields::FieldTC>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 3\" \\n \\f B \\l 1", System::String::Empty, fieldTc);
    ASSERT_TRUE(fieldTc->get_OmitPageNumber());
    ASSERT_EQ(fieldTc->get_Text(), u"TC field 3");
    ASSERT_EQ(fieldTc->get_TypeIdentifier(), u"B");
    ASSERT_EQ(fieldTc->get_EntryLevel(), u"1");
    
    fieldTc = System::ExplicitCast<Aspose::Words::Fields::FieldTC>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldTOCEntry, u" TC  \"TC field 4\" \\n \\f A \\l 5", System::String::Empty, fieldTc);
    ASSERT_TRUE(fieldTc->get_OmitPageNumber());
    ASSERT_EQ(fieldTc->get_Text(), u"TC field 4");
    ASSERT_EQ(fieldTc->get_TypeIdentifier(), u"A");
    ASSERT_EQ(fieldTc->get_EntryLevel(), u"5");
}

System::SharedPtr<Aspose::Words::Fields::FieldIncludeText> ExField::CreateFieldIncludeText(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String sourceFullName, bool lockFields, System::String mimeType, System::String textConverter, System::String encoding)
{
    auto fieldIncludeText = System::ExplicitCast<Aspose::Words::Fields::FieldIncludeText>(builder->InsertField(FieldType::FieldIncludeText, true));
    fieldIncludeText->set_SourceFullName(sourceFullName);
    fieldIncludeText->set_LockFields(lockFields);
    fieldIncludeText->set_MimeType(mimeType);
    fieldIncludeText->set_TextConverter(textConverter);
    fieldIncludeText->set_Encoding(encoding);
    
    return fieldIncludeText;
}

void ExField::TestFieldIncludeText(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    
    auto fieldIncludeText = System::ExplicitCast<Aspose::Words::Fields::FieldIncludeText>(doc->get_Range()->get_Fields()->idx_get(0));
    ASSERT_EQ(fieldIncludeText->get_SourceFullName(), get_MyDir() + u"CD collection data.xml");
    ASSERT_EQ(fieldIncludeText->get_XslTransformation(), get_MyDir() + u"CD collection XSL transformation.xsl");
    ASSERT_FALSE(fieldIncludeText->get_LockFields());
    ASSERT_EQ(fieldIncludeText->get_MimeType(), u"text/xml");
    ASSERT_EQ(fieldIncludeText->get_TextConverter(), u"XML");
    ASSERT_EQ(fieldIncludeText->get_Encoding(), u"ISO-8859-1");
    ASSERT_EQ(fieldIncludeText->GetFieldCode(), System::String(u" INCLUDETEXT  \"") + get_MyDir().Replace(u"\\", u"\\\\") + u"CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\t \"" + get_MyDir().Replace(u"\\", u"\\\\") + u"CD collection XSL transformation.xsl\"");
    ASSERT_TRUE(fieldIncludeText->get_Result().StartsWith(u"My CD Collection"));
    
    auto cdCollectionData = System::MakeObject<System::Xml::XmlDocument>();
    cdCollectionData->LoadXml(System::IO::File::ReadAllText(get_MyDir() + u"CD collection data.xml"));
    System::SharedPtr<System::Xml::XmlNode> catalogData = cdCollectionData->get_ChildNodes()->idx_get(0);
    
    auto cdCollectionXslTransformation = System::MakeObject<System::Xml::XmlDocument>();
    cdCollectionXslTransformation->LoadXml(System::IO::File::ReadAllText(get_MyDir() + u"CD collection XSL transformation.xsl"));
    
    System::SharedPtr<Aspose::Words::Tables::Table> table = doc->get_FirstSection()->get_Body()->get_Tables()->idx_get(0);
    
    auto manager = System::MakeObject<System::Xml::XmlNamespaceManager>(cdCollectionXslTransformation->get_NameTable());
    manager->AddNamespace(u"xsl", u"http://www.w3.org/1999/XSL/Transform");
    
    for (int32_t i = 0; i < table->get_Rows()->get_Count(); i++)
    {
        for (int32_t j = 0; j < table->get_Rows()->idx_get(i)->get_Count(); j++)
        {
            if (i == 0)
            {
                // When on the first row from the input document's table, ensure that all table's cells match all XML element Names.
                for (int32_t k = 0; k < table->get_Rows()->get_Count() - 1; k++)
                {
                    ASSERT_EQ(table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->GetText().Replace(ControlChar::Cell(), System::String::Empty).ToLower(), catalogData->get_ChildNodes()->idx_get(k)->get_ChildNodes()->idx_get(j)->get_Name());
                }
                
                // Also, make sure that the whole first row has the same color as the XSL transform.
                ASSERT_EQ(System::Drawing::ColorTranslator::ToHtml(table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->get_CellFormat()->get_Shading()->get_BackgroundPatternColor()).ToLower(), cdCollectionXslTransformation->SelectNodes(u"//xsl:stylesheet/xsl:template/html/body/table/tr", manager)->idx_get(0)->get_Attributes()->GetNamedItem(u"bgcolor")->get_Value());
            }
            else
            {
                // When on all other rows of the input document's table, ensure that cell contents match XML element Values.
                ASSERT_EQ(table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->GetText().Replace(ControlChar::Cell(), System::String::Empty), catalogData->get_ChildNodes()->idx_get(i - 1)->get_ChildNodes()->idx_get(j)->get_FirstChild()->get_Value());
                ASPOSE_ASSERT_EQ(table->get_Rows()->idx_get(i)->get_Cells()->idx_get(j)->get_CellFormat()->get_Shading()->get_BackgroundPatternColor(), System::Drawing::Color::Empty);
            }
            
            ASPOSE_ASSERT_EQ(table->get_FirstRow()->get_RowFormat()->get_Borders()->get_Bottom()->get_LineWidth(), System::Double::Parse(cdCollectionXslTransformation->SelectNodes(u"//xsl:stylesheet/xsl:template/html/body/table", manager)->idx_get(0)->get_Attributes()->GetNamedItem(u"border")->get_Value()) * 0.75);
        }
    }
    
    fieldIncludeText = System::ExplicitCast<Aspose::Words::Fields::FieldIncludeText>(doc->get_Range()->get_Fields()->idx_get(1));
    ASSERT_EQ(fieldIncludeText->get_SourceFullName(), get_MyDir() + u"CD collection data.xml");
    ASSERT_TRUE(System::TestTools::IsNull(fieldIncludeText->get_XslTransformation()));
    ASSERT_FALSE(fieldIncludeText->get_LockFields());
    ASSERT_EQ(fieldIncludeText->get_MimeType(), u"text/xml");
    ASSERT_EQ(fieldIncludeText->get_TextConverter(), u"XML");
    ASSERT_EQ(fieldIncludeText->get_Encoding(), u"ISO-8859-1");
    ASSERT_EQ(fieldIncludeText->GetFieldCode(), System::String(u" INCLUDETEXT  \"") + get_MyDir().Replace(u"\\", u"\\\\") + u"CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\n xmlns:n='myNamespace' \\x /catalog/cd/title");
    
    System::String expectedFieldResult = u"";
    for (int32_t i = 0; i < catalogData->get_ChildNodes()->get_Count(); i++)
    {
        expectedFieldResult += catalogData->get_ChildNodes()->idx_get(i)->get_ChildNodes()->idx_get(0)->get_ChildNodes()->idx_get(0)->get_Value();
    }
    
    ASSERT_EQ(fieldIncludeText->get_Result(), expectedFieldResult);
}

void ExField::TestMergeFieldImageDimension(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 0);
    ASSERT_EQ(doc->GetChildNodes(NodeType::Shape, true)->get_Count(), 3);
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(NodeType::Shape, 0, true));
    
    TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
    ASPOSE_ASSERT_EQ(shape->get_Width(), 200.0);
    ASPOSE_ASSERT_EQ(shape->get_Height(), 200.0);
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(NodeType::Shape, 1, true));
    
    TestUtil::VerifyImageInShape(400, 400, ImageType::Png, shape);
    ASPOSE_ASSERT_EQ(shape->get_Width(), 200.0);
    ASPOSE_ASSERT_EQ(shape->get_Height(), 200.0);
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(NodeType::Shape, 2, true));
    
    TestUtil::VerifyImageInShape(534, 534, ImageType::Emf, shape);
    ASPOSE_ASSERT_EQ(shape->get_Width(), 200.0);
    ASPOSE_ASSERT_EQ(shape->get_Height(), 200.0);
}

void ExField::TestMergeFieldImages(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 0);
    ASSERT_EQ(doc->GetChildNodes(NodeType::Shape, true)->get_Count(), 2);
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(NodeType::Shape, 0, true));
    
    TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, shape);
    ASPOSE_ASSERT_EQ(shape->get_Width(), 300.0);
    ASPOSE_ASSERT_EQ(shape->get_Height(), 300.0);
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(NodeType::Shape, 1, true));
    
    TestUtil::VerifyImageInShape(400, 400, ImageType::Png, shape);
    ASSERT_NEAR(shape->get_Width(), 300.0, 1);
    ASSERT_NEAR(shape->get_Height(), 300.0, 1);
}

void ExField::InsertFieldLink(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::ApiExamples::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool shouldAutoUpdate)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldLink>(builder->InsertField(FieldType::FieldLink, true));
    
    switch (insertLinkedObjectAs)
    {
        case InsertLinkedObjectAs::Text:
            field->set_InsertAsText(true);
            break;
        case InsertLinkedObjectAs::Unicode:
            field->set_InsertAsUnicode(true);
            break;
        case InsertLinkedObjectAs::Html:
            field->set_InsertAsHtml(true);
            break;
        case InsertLinkedObjectAs::Rtf:
            field->set_InsertAsRtf(true);
            break;
        case InsertLinkedObjectAs::Picture:
            field->set_InsertAsPicture(true);
            break;
        case InsertLinkedObjectAs::Bitmap:
            field->set_InsertAsBitmap(true);
            break;
    }
    
    field->set_AutoUpdate(shouldAutoUpdate);
    field->set_ProgId(progId);
    field->set_SourceFullName(sourceFullName);
    field->set_SourceItem(sourceItem);
    
    builder->Writeln(u"\n");
}

void ExField::InsertFieldDde(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::ApiExamples::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool isLinked, bool shouldAutoUpdate)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldDde>(builder->InsertField(FieldType::FieldDDE, true));
    
    switch (insertLinkedObjectAs)
    {
        case InsertLinkedObjectAs::Text:
            field->set_InsertAsText(true);
            break;
        case InsertLinkedObjectAs::Unicode:
            field->set_InsertAsUnicode(true);
            break;
        case InsertLinkedObjectAs::Html:
            field->set_InsertAsHtml(true);
            break;
        case InsertLinkedObjectAs::Rtf:
            field->set_InsertAsRtf(true);
            break;
        case InsertLinkedObjectAs::Picture:
            field->set_InsertAsPicture(true);
            break;
        case InsertLinkedObjectAs::Bitmap:
            field->set_InsertAsBitmap(true);
            break;
    }
    
    field->set_AutoUpdate(shouldAutoUpdate);
    field->set_ProgId(progId);
    field->set_SourceFullName(sourceFullName);
    field->set_SourceItem(sourceItem);
    field->set_IsLinked(isLinked);
    
    builder->Writeln(u"\n");
}

void ExField::InsertFieldDdeAuto(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::ApiExamples::InsertLinkedObjectAs insertLinkedObjectAs, System::String progId, System::String sourceFullName, System::String sourceItem, bool isLinked)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldDdeAuto>(builder->InsertField(FieldType::FieldDDEAuto, true));
    
    switch (insertLinkedObjectAs)
    {
        case InsertLinkedObjectAs::Text:
            field->set_InsertAsText(true);
            break;
        case InsertLinkedObjectAs::Unicode:
            field->set_InsertAsUnicode(true);
            break;
        case InsertLinkedObjectAs::Html:
            field->set_InsertAsHtml(true);
            break;
        case InsertLinkedObjectAs::Rtf:
            field->set_InsertAsRtf(true);
            break;
        case InsertLinkedObjectAs::Picture:
            field->set_InsertAsPicture(true);
            break;
        case InsertLinkedObjectAs::Bitmap:
            field->set_InsertAsBitmap(true);
            break;
    }
    
    field->set_ProgId(progId);
    field->set_SourceFullName(sourceFullName);
    field->set_SourceItem(sourceItem);
    field->set_IsLinked(isLinked);
}

void ExField::TestFieldFillIn(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 1);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldFillIn>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldFillIn, u" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", u"Response modified by PromptRespondent. A default response.", field);
    ASSERT_EQ(field->get_PromptText(), u"Please enter a response:");
    ASSERT_EQ(field->get_DefaultResponse(), u"A default response.");
    ASSERT_TRUE(field->get_PromptOnceOnMailMerge());
}

void ExField::InsertMergeFields(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String firstFieldTextBefore)
{
    InsertMergeField(builder, u"Courtesy Title", firstFieldTextBefore, u" ");
    InsertMergeField(builder, u"First Name", nullptr, u" ");
    InsertMergeField(builder, u"Last Name", nullptr, nullptr);
    builder->InsertParagraph();
}

void ExField::InsertMergeField(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String fieldName, System::String textBefore, System::String textAfter)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldMergeField>(builder->InsertField(FieldType::FieldMergeField, true));
    field->set_FieldName(fieldName);
    field->set_TextBefore(textBefore);
    field->set_TextAfter(textAfter);
}

void ExField::TestFieldNext(System::SharedPtr<Aspose::Words::Document> doc)
{
    doc = DocumentHelper::SaveOpen(doc);
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 0);
    ASSERT_EQ(doc->GetText(), System::String(u"First row: Mr. John Doe\r") + u"Second row: Mrs. Jane Cardholder\r" + u"Third row: Mr. Joe Bloggs\r\f");
}

System::SharedPtr<Aspose::Words::Fields::FieldNoteRef> ExField::InsertFieldNoteRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, bool insertReferenceMark, System::String textBefore)
{
    builder->Write(textBefore);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldNoteRef>(builder->InsertField(FieldType::FieldNoteRef, true));
    field->set_BookmarkName(bookmarkName);
    field->set_InsertHyperlink(insertHyperlink);
    field->set_InsertRelativePosition(insertRelativePosition);
    field->set_InsertReferenceMark(insertReferenceMark);
    builder->Writeln();
    
    return field;
}

void ExField::InsertBookmarkWithFootnote(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String bookmarkText, System::String footnoteText)
{
    builder->StartBookmark(bookmarkName);
    builder->Write(bookmarkText);
    builder->InsertFootnote(FootnoteType::Footnote, footnoteText);
    builder->EndBookmark(bookmarkName);
    builder->Writeln();
}

void ExField::TestNoteRef(System::SharedPtr<Aspose::Words::Document> doc)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldNoteRef>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldNoteRef, u" NOTEREF  MyBookmark2 \\h", u"2", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark2");
    ASSERT_TRUE(field->get_InsertHyperlink());
    ASSERT_FALSE(field->get_InsertRelativePosition());
    ASSERT_FALSE(field->get_InsertReferenceMark());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldNoteRef>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldNoteRef, u" NOTEREF  MyBookmark1 \\h \\p", u"1 above", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark1");
    ASSERT_TRUE(field->get_InsertHyperlink());
    ASSERT_TRUE(field->get_InsertRelativePosition());
    ASSERT_FALSE(field->get_InsertReferenceMark());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldNoteRef>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldNoteRef, u" NOTEREF  MyBookmark2 \\h \\p \\f", u"2 below", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark2");
    ASSERT_TRUE(field->get_InsertHyperlink());
    ASSERT_TRUE(field->get_InsertRelativePosition());
    ASSERT_TRUE(field->get_InsertReferenceMark());
}

System::SharedPtr<Aspose::Words::Fields::FieldPageRef> ExField::InsertFieldPageRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, bool insertHyperlink, bool insertRelativePosition, System::String textBefore)
{
    builder->Write(textBefore);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(builder->InsertField(FieldType::FieldPageRef, true));
    field->set_BookmarkName(bookmarkName);
    field->set_InsertHyperlink(insertHyperlink);
    field->set_InsertRelativePosition(insertRelativePosition);
    builder->Writeln();
    
    return field;
}

void ExField::InsertAndNameBookmark(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName)
{
    builder->StartBookmark(bookmarkName);
    builder->Writeln(System::String::Format(u"Contents of bookmark \"{0}\".", bookmarkName));
    builder->EndBookmark(bookmarkName);
}

void ExField::TestPageRef(System::SharedPtr<Aspose::Words::Document> doc)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark3 \\h", u"2", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark3");
    ASSERT_TRUE(field->get_InsertHyperlink());
    ASSERT_FALSE(field->get_InsertRelativePosition());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark1 \\h \\p", u"above", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark1");
    ASSERT_TRUE(field->get_InsertHyperlink());
    ASSERT_TRUE(field->get_InsertRelativePosition());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark2 \\h \\p", u"below", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark2");
    ASSERT_TRUE(field->get_InsertHyperlink());
    ASSERT_TRUE(field->get_InsertRelativePosition());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF  MyBookmark3 \\h \\p", u"on page 2", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark3");
    ASSERT_TRUE(field->get_InsertHyperlink());
    ASSERT_TRUE(field->get_InsertRelativePosition());
}

System::SharedPtr<Aspose::Words::Fields::FieldRef> ExField::InsertFieldRef(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String bookmarkName, System::String textBefore, System::String textAfter)
{
    builder->Write(textBefore);
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(builder->InsertField(FieldType::FieldRef, true));
    field->set_BookmarkName(bookmarkName);
    builder->Write(textAfter);
    return field;
}

void ExField::TestFieldRef(System::SharedPtr<Aspose::Words::Document> doc)
{
    TestUtil::VerifyFootnote(FootnoteType::Footnote, true, System::String::Empty, u"MyBookmark footnote #1", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
    TestUtil::VerifyFootnote(FootnoteType::Footnote, true, System::String::Empty, u"MyBookmark footnote #2", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(NodeType::Footnote, 1, true)));
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\f \\h", u"Text that will appear in REF field", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark");
    ASSERT_TRUE(field->get_IncludeNoteOrComment());
    ASSERT_TRUE(field->get_InsertHyperlink());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\p", u"below", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark");
    ASSERT_TRUE(field->get_InsertRelativePosition());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\n", u"‎>>> i", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark");
    ASSERT_TRUE(field->get_InsertParagraphNumber());
    ASSERT_EQ(field->GetFieldCode(), u" REF  MyBookmark \\n");
    ASSERT_EQ(field->get_Result(), u"‎>>> i");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\n \\t", u"‎i", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark");
    ASSERT_TRUE(field->get_InsertParagraphNumber());
    ASSERT_TRUE(field->get_SuppressNonDelimiters());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\w", u"‎> 4>> c>>> i", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark");
    ASSERT_TRUE(field->get_InsertParagraphNumberInFullContext());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(doc->get_Range()->get_Fields()->idx_get(5));
    
    TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark \\r", u"‎>> c>>> i", field);
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark");
    ASSERT_TRUE(field->get_InsertParagraphNumberInRelativeContext());
}

System::SharedPtr<Aspose::Words::Fields::FieldTA> ExField::InsertToaEntry(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String entryCategory, System::String longCitation)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldTA>(builder->InsertField(FieldType::FieldTOAEntry, false));
    field->set_EntryCategory(entryCategory);
    field->set_LongCitation(longCitation);
    
    builder->InsertBreak(BreakType::PageBreak);
    
    return field;
}

void ExField::TestFieldTOA(System::SharedPtr<Aspose::Words::Document> doc)
{
    auto fieldTOA = System::ExplicitCast<Aspose::Words::Fields::FieldToa>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(fieldTOA->get_EntryCategory(), u"1");
    ASSERT_TRUE(fieldTOA->get_UseHeading());
    ASSERT_EQ(fieldTOA->get_BookmarkName(), u"MyBookmark");
    ASSERT_EQ(fieldTOA->get_EntrySeparator(), u" \t p.");
    ASSERT_EQ(fieldTOA->get_PageNumberListSeparator(), u" & p. ");
    ASSERT_TRUE(fieldTOA->get_UsePassim());
    ASSERT_EQ(fieldTOA->get_PageRangeSeparator(), u" to ");
    ASSERT_TRUE(fieldTOA->get_RemoveEntryFormatting());
    ASSERT_EQ(fieldTOA->GetFieldCode(), u" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f");
    ASSERT_EQ(fieldTOA->get_Result(), System::String(u"Cases\r") + u"Source 2 \t p.5\r" + u"Source 3 \t p.4 & p. 7 to 10\r" + u"Source 4 \t p.passim\r");
    
    auto fieldTA = System::ExplicitCast<Aspose::Words::Fields::FieldTA>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 1\"", System::String::Empty, fieldTA);
    ASSERT_EQ(fieldTA->get_EntryCategory(), u"1");
    ASSERT_EQ(fieldTA->get_LongCitation(), u"Source 1");
    
    fieldTA = System::ExplicitCast<Aspose::Words::Fields::FieldTA>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 2 \\l \"Source 2\"", System::String::Empty, fieldTA);
    ASSERT_EQ(fieldTA->get_EntryCategory(), u"2");
    ASSERT_EQ(fieldTA->get_LongCitation(), u"Source 2");
    
    fieldTA = System::ExplicitCast<Aspose::Words::Fields::FieldTA>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 3\" \\s S.3", System::String::Empty, fieldTA);
    ASSERT_EQ(fieldTA->get_EntryCategory(), u"1");
    ASSERT_EQ(fieldTA->get_LongCitation(), u"Source 3");
    ASSERT_EQ(fieldTA->get_ShortCitation(), u"S.3");
    
    fieldTA = System::ExplicitCast<Aspose::Words::Fields::FieldTA>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 2\" \\b \\i", System::String::Empty, fieldTA);
    ASSERT_EQ(fieldTA->get_EntryCategory(), u"1");
    ASSERT_EQ(fieldTA->get_LongCitation(), u"Source 2");
    ASSERT_TRUE(fieldTA->get_IsBold());
    ASSERT_TRUE(fieldTA->get_IsItalic());
    
    fieldTA = System::ExplicitCast<Aspose::Words::Fields::FieldTA>(doc->get_Range()->get_Fields()->idx_get(5));
    
    TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", System::String::Empty, fieldTA);
    ASSERT_EQ(fieldTA->get_EntryCategory(), u"1");
    ASSERT_EQ(fieldTA->get_LongCitation(), u"Source 3");
    ASSERT_EQ(fieldTA->get_PageRangeBookmarkName(), u"MyMultiPageBookmark");
    
    for (int32_t i = 6; i < 11; i++)
    {
        fieldTA = System::ExplicitCast<Aspose::Words::Fields::FieldTA>(doc->get_Range()->get_Fields()->idx_get(i));
        
        TestUtil::VerifyField(FieldType::FieldTOAEntry, u" TA  \\c 1 \\l \"Source 4\"", System::String::Empty, fieldTA);
        ASSERT_EQ(fieldTA->get_EntryCategory(), u"1");
        ASSERT_EQ(fieldTA->get_LongCitation(), u"Source 4");
    }
}

System::SharedPtr<Aspose::Words::Fields::FieldEQ> ExField::InsertFieldEQ(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String args)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldEQ>(builder->InsertField(FieldType::FieldEquation, true));
    builder->MoveTo(field->get_Separator());
    builder->Write(args);
    builder->MoveTo(field->get_Start()->get_ParentNode());
    
    builder->InsertParagraph();
    return field;
}

void ExField::TestFieldEQ(System::SharedPtr<Aspose::Words::Document> doc)
{
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\f(1,4)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(1));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(2));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ A \\d \\fo30 \\li() B", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(3));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(4));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\i \\su(n=1,5,n)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(5));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\l(1,1,2,3,n,8,13)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(6));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\r (3,x)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(7));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\s \\up8(Superscript) Text \\s \\do8(Subscript)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(8));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\x \\to \\bo \\le \\ri(5)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(9));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(10));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(11));
    TestUtil::VerifyField(FieldType::FieldEquation, u" EQ \\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(12));
}

System::SharedPtr<Aspose::Words::Fields::FieldTime> ExField::InsertFieldTime(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::String format)
{
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldTime>(builder->InsertField(FieldType::FieldTime, true));
    builder->MoveTo(field->get_Separator());
    builder->Write(format);
    builder->MoveTo(field->get_Start()->get_ParentNode());
    
    builder->InsertParagraph();
    return field;
}

void ExField::TestFieldTime(System::SharedPtr<Aspose::Words::Document> doc)
{
    System::DateTime docLoadingTime = System::DateTime::get_Now();
    doc = DocumentHelper::SaveOpen(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldTime>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->GetFieldCode(), u" TIME ");
    ASSERT_EQ(field->get_Type(), FieldType::FieldTime);
    ASSERT_EQ(System::DateTime::get_Today().AddHours(docLoadingTime.get_Hour()).AddMinutes(docLoadingTime.get_Minute()), System::DateTime::Parse(field->get_Result()));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTime>(doc->get_Range()->get_Fields()->idx_get(1));
    
    ASSERT_EQ(field->GetFieldCode(), u" TIME \\@ HHmm");
    ASSERT_EQ(field->get_Type(), FieldType::FieldTime);
    ASSERT_EQ(System::DateTime::get_Today().AddHours(docLoadingTime.get_Hour()).AddMinutes(docLoadingTime.get_Minute()), System::DateTime::Parse(field->get_Result()));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTime>(doc->get_Range()->get_Fields()->idx_get(2));
    
    ASSERT_EQ(field->GetFieldCode(), u" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"");
    ASSERT_EQ(field->get_Type(), FieldType::FieldTime);
    ASSERT_EQ(System::DateTime::get_Today().AddHours(docLoadingTime.get_Hour()).AddMinutes(docLoadingTime.get_Minute()), System::DateTime::Parse(field->get_Result()));
}

namespace gtest_test
{

class ExField : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExField> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExField>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExField> ExField::s_instance;

} // namespace gtest_test

void ExField::GetFieldFromDocument()
{
    //ExStart
    //ExFor:FieldType
    //ExFor:FieldChar
    //ExFor:FieldChar.FieldType
    //ExFor:FieldChar.IsDirty
    //ExFor:FieldChar.IsLocked
    //ExFor:FieldChar.GetField
    //ExFor:Field.IsLocked
    //ExSummary:Shows how to work with a FieldStart node.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(builder->InsertField(FieldType::FieldDate, true));
    field->get_Format()->set_DateTimeFormat(u"dddd, MMMM dd, yyyy");
    field->Update();
    
    System::SharedPtr<Aspose::Words::Fields::FieldChar> fieldStart = field->get_Start();
    
    ASSERT_EQ(fieldStart->get_FieldType(), FieldType::FieldDate);
    ASPOSE_ASSERT_EQ(fieldStart->get_IsDirty(), false);
    ASPOSE_ASSERT_EQ(fieldStart->get_IsLocked(), false);
    
    // Retrieve the facade object which represents the field in the document.
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(fieldStart->GetField());
    
    ASPOSE_ASSERT_EQ(field->get_IsLocked(), false);
    ASSERT_EQ(field->GetFieldCode(), (u" DATE  \\@ \"dddd, MMMM dd, yyyy\""));
    
    // Update the field to show the current date.
    field->Update();
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    
    TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\@ \"dddd, MMMM dd, yyyy\"", System::DateTime::get_Now().ToString(u"dddd, MMMM dd, yyyy"), doc->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, GetFieldFromDocument)
{
    s_instance->GetFieldFromDocument();
}

} // namespace gtest_test

void ExField::GetFieldData()
{
    //ExStart
    //ExFor:FieldStart.FieldData
    //ExSummary:Shows how to get data associated with the field.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Field sample - Field with data.docx"));
    
    System::SharedPtr<Aspose::Words::Fields::Field> field = doc->get_Range()->get_Fields()->idx_get(2);
    std::cout << System::Text::Encoding::get_Default()->GetString(field->get_Start()->get_FieldData()) << std::endl;
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, GetFieldData)
{
    s_instance->GetFieldData();
}

} // namespace gtest_test

void ExField::GetFieldCode()
{
    //ExStart
    //ExFor:Field.GetFieldCode
    //ExFor:Field.GetFieldCode(bool)
    //ExSummary:Shows how to get a field's field code.
    // Open a document which contains a MERGEFIELD inside an IF field.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Nested fields.docx"));
    auto fieldIf = System::ExplicitCast<Aspose::Words::Fields::FieldIf>(doc->get_Range()->get_Fields()->idx_get(0));
    
    // There are two ways of getting a field's field code:
    // 1 -  Omit its inner fields:
    ASSERT_EQ(fieldIf->GetFieldCode(false), u" IF  > 0 \" (surplus of ) \" \"\" ");
    
    // 2 -  Include its inner fields:
    ASSERT_EQ(fieldIf->GetFieldCode(true), System::String::Format(u" IF \u0013 MERGEFIELD NetIncome \u0014\u0015 > 0 \" (surplus of \u0013 MERGEFIELD  NetIncome \\f $ \u0014\u0015) \" \"\" ", System::MakeObject<System::Array<System::SharedPtr<System::Object>>>(0)));
    
    // By default, the GetFieldCode method displays inner fields.
    ASSERT_EQ(fieldIf->GetFieldCode(true), fieldIf->GetFieldCode());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, GetFieldCode)
{
    s_instance->GetFieldCode();
}

} // namespace gtest_test

void ExField::DisplayResult()
{
    //ExStart
    //ExFor:Field.DisplayResult
    //ExSummary:Shows how to get the real text that a field displays in the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"This document was written by ");
    auto fieldAuthor = System::ExplicitCast<Aspose::Words::Fields::FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));
    fieldAuthor->set_AuthorName(u"John Doe");
    
    // We can use the DisplayResult property to verify what exact text
    // a field would display in its place in the document.
    ASSERT_EQ(fieldAuthor->get_DisplayResult(), System::String::Empty);
    
    // Fields do not maintain accurate result values in real-time.
    // To make sure our fields display accurate results at any given time,
    // such as right before a save operation, we need to update them manually.
    fieldAuthor->Update();
    
    ASSERT_EQ(fieldAuthor->get_DisplayResult(), u"John Doe");
    
    doc->Save(get_ArtifactsDir() + u"Field.DisplayResult.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.DisplayResult.docx"));
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(0)->get_DisplayResult(), u"John Doe");
}

namespace gtest_test
{

TEST_F(ExField, DisplayResult)
{
    s_instance->DisplayResult();
}

} // namespace gtest_test

void ExField::CreateWithFieldBuilder()
{
    //ExStart
    //ExFor:FieldBuilder.#ctor(FieldType)
    //ExFor:FieldBuilder.BuildAndInsert(Inline)
    //ExSummary:Shows how to create and insert a field using a field builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A convenient way of adding text content to a document is with a document builder.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u" Hello world! This text is one Run, which is an inline node.");
    
    // Fields have their builder, which we can use to construct a field code piece by piece.
    // In this case, we will construct a BARCODE field representing a US postal code,
    // and then insert it in front of a Run.
    auto fieldBuilder = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldBarcode);
    fieldBuilder->AddArgument(u"90210");
    fieldBuilder->AddSwitch(u"\\f", u"A");
    fieldBuilder->AddSwitch(u"\\u");
    
    fieldBuilder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0));
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.CreateWithFieldBuilder.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.CreateWithFieldBuilder.docx"));
    
    TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE 90210 \\f A \\u ", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
    
    ASPOSE_ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(0)->get_End(), doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(11)->get_PreviousSibling());
    ASSERT_EQ(doc->GetText().Trim(), System::String::Format(u"{0} BARCODE 90210 \\f A \\u {1} Hello world! This text is one Run, which is an inline node.", ControlChar::FieldStartChar, ControlChar::FieldEndChar));
}

namespace gtest_test
{

TEST_F(ExField, CreateWithFieldBuilder)
{
    s_instance->CreateWithFieldBuilder();
}

} // namespace gtest_test

void ExField::RevNum()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.RevisionNumber
    //ExFor:FieldRevNum
    //ExSummary:Shows how to work with REVNUM fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Current revision #");
    
    // Insert a REVNUM field, which displays the document's current revision number property.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldRevNum>(builder->InsertField(FieldType::FieldRevisionNum, true));
    
    ASSERT_EQ(field->GetFieldCode(), u" REVNUM ");
    ASSERT_EQ(field->get_Result(), u"1");
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_RevisionNumber(), 1);
    
    // This property counts how many times a document has been saved in Microsoft Word,
    // and is unrelated to tracked revisions. We can find it by right clicking the document in Windows Explorer
    // via Properties -> Details. We can update this property manually.
    System::WithLambda::setter_post_increment_wrap(GETTER_SETTER_LVAL_LAMBDA_ARGS(doc->get_BuiltInDocumentProperties(), RevisionNumber));
    ASSERT_EQ(field->get_Result(), u"1");
    //ExSkip
    field->Update();
    
    ASSERT_EQ(field->get_Result(), u"2");
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_RevisionNumber(), 2);
    
    TestUtil::VerifyField(FieldType::FieldRevisionNum, u" REVNUM ", u"2", doc->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, RevNum)
{
    s_instance->RevNum();
}

} // namespace gtest_test

void ExField::InsertFieldNone()
{
    //ExStart
    //ExFor:FieldUnknown
    //ExSummary:Shows how to work with 'FieldNone' field in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a field that does not denote an objective field type in its field code.
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->InsertField(u" NOTAREALFIELD //a");
    
    // The "FieldNone" field type is reserved for fields such as these.
    ASSERT_EQ(field->get_Type(), FieldType::FieldNone);
    
    // We can also still work with these fields and assign them as instances of the FieldUnknown class.
    auto fieldUnknown = System::ExplicitCast<Aspose::Words::Fields::FieldUnknown>(field);
    ASSERT_EQ(fieldUnknown->GetFieldCode(), u" NOTAREALFIELD //a");
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    
    TestUtil::VerifyField(FieldType::FieldNone, u" NOTAREALFIELD //a", u"Error! Bookmark not defined.", doc->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, InsertFieldNone)
{
    s_instance->InsertFieldNone();
}

} // namespace gtest_test

void ExField::InsertTcField()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a TC field at the current document builder position.
    builder->InsertField(u"TC \"Entry Text\" \\f t");
}

namespace gtest_test
{

TEST_F(ExField, InsertTcField)
{
    s_instance->InsertTcField();
}

} // namespace gtest_test

void ExField::InsertTcFieldsAtText()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto options = System::MakeObject<Aspose::Words::Replacing::FindReplaceOptions>();
    options->set_ReplacingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExField::InsertTcFieldHandler>(u"Chapter 1", u"\\l 1"));
    
    // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
    doc->get_Range()->Replace(System::MakeObject<System::Text::RegularExpressions::Regex>(u"The Beginning"), u"", options);
}

namespace gtest_test
{

TEST_F(ExField, InsertTcFieldsAtText)
{
    s_instance->InsertTcFieldsAtText();
}

} // namespace gtest_test

void ExField::FieldLocale()
{
    //ExStart
    //ExFor:Field.LocaleId
    //ExSummary:Shows how to insert a field and work with its locale.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a DATE field, and then print the date it will display.
    // Your thread's current culture determines the formatting of the date.
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->InsertField(u"DATE");
    std::cout << System::String::Format(u"Today's date, as displayed in the \"{0}\" culture: {1}", System::Globalization::CultureInfo::get_CurrentCulture()->get_EnglishName(), field->get_Result()) << std::endl;
    
    ASSERT_EQ(field->get_LocaleId(), 1033);
    ASSERT_EQ(doc->get_FieldOptions()->get_FieldUpdateCultureSource(), FieldUpdateCultureSource::CurrentThread);
    //ExSkip
    
    // Changing the culture of our thread will impact the result of the DATE field.
    // Another way to get the DATE field to display a date in a different culture is to use its LocaleId property.
    // This way allows us to avoid changing the thread's culture to get this effect.
    doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
    auto de = System::MakeObject<System::Globalization::CultureInfo>(u"de-DE");
    field->set_LocaleId(de->get_LCID());
    field->Update();
    
    std::cout << System::String::Format(u"Today's date, as displayed according to the \"{0}\" culture: {1}", System::Globalization::CultureInfo::GetCultureInfo(field->get_LocaleId())->get_EnglishName(), field->get_Result()) << std::endl;
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    field = doc->get_Range()->get_Fields()->idx_get(0);
    
    TestUtil::VerifyField(FieldType::FieldDate, u"DATE", System::DateTime::get_Now().ToString(de->get_DateTimeFormat()->get_ShortDatePattern()), field);
    ASSERT_EQ(field->get_LocaleId(), System::MakeObject<System::Globalization::CultureInfo>(u"de-DE")->get_LCID());
}

namespace gtest_test
{

TEST_F(ExField, FieldLocale)
{
    s_instance->FieldLocale();
}

} // namespace gtest_test

void ExField::UpdateDirtyFields(bool updateDirtyFields)
{
    //ExStart
    //ExFor:Field.IsDirty
    //ExFor:LoadOptions.UpdateDirtyFields
    //ExSummary:Shows how to use special property for updating field result.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Give the document's built-in "Author" property value, and then display it with a field.
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));
    
    ASSERT_FALSE(field->get_IsDirty());
    ASSERT_EQ(field->get_Result(), u"John Doe");
    
    // Update the property. The field still displays the old value.
    doc->get_BuiltInDocumentProperties()->set_Author(u"John & Jane Doe");
    
    ASSERT_EQ(field->get_Result(), u"John Doe");
    
    // Since the field's value is out of date, we can mark it as "dirty".
    // This value will stay out of date until we update the field manually with the Field.Update() method.
    field->set_IsDirty(true);
    
    {
        auto docStream = System::MakeObject<System::IO::MemoryStream>();
        // If we save without calling an update method,
        // the field will keep displaying the out of date value in the output document.
        doc->Save(docStream, SaveFormat::Docx);
        
        // The LoadOptions object has an option to update all fields
        // marked as "dirty" when loading the document.
        auto options = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
        options->set_UpdateDirtyFields(updateDirtyFields);
        doc = System::MakeObject<Aspose::Words::Document>(docStream, options);
        
        ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Author(), u"John & Jane Doe");
        
        field = System::ExplicitCast<Aspose::Words::Fields::FieldAuthor>(doc->get_Range()->get_Fields()->idx_get(0));
        
        // Updating dirty fields like this automatically set their "IsDirty" flag to false.
        if (updateDirtyFields)
        {
            ASSERT_EQ(field->get_Result(), u"John & Jane Doe");
            ASSERT_FALSE(field->get_IsDirty());
        }
        else
        {
            ASSERT_EQ(field->get_Result(), u"John Doe");
            ASSERT_TRUE(field->get_IsDirty());
        }
    }
    //ExEnd
}

namespace gtest_test
{

using ExField_UpdateDirtyFields_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExField::UpdateDirtyFields)>::type;

struct ExField_UpdateDirtyFields : public ExField, public Aspose::Words::ApiExamples::ExField, public ::testing::WithParamInterface<ExField_UpdateDirtyFields_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExField_UpdateDirtyFields, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateDirtyFields(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_UpdateDirtyFields, ::testing::ValuesIn(ExField_UpdateDirtyFields::TestCases()));

} // namespace gtest_test

void ExField::InsertFieldWithFieldBuilderException()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Run> run = DocumentHelper::InsertNewRun(doc, u" Hello World!", 0);
    
    auto argumentBuilder = System::MakeObject<Aspose::Words::Fields::FieldArgumentBuilder>();
    argumentBuilder->AddField(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldMergeField));
    argumentBuilder->AddNode(run);
    argumentBuilder->AddText(u"Text argument builder");
    
    auto fieldBuilder = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIncludeText);
    
    ASSERT_THROW(System::Action<>(static_cast<std::function<void()>>([&fieldBuilder, &argumentBuilder, &run]() -> void
    {
        fieldBuilder->AddArgument(argumentBuilder)->AddArgument(u"=")->AddArgument(u"BestField")->AddArgument(10)->AddArgument(20.0)->BuildAndInsert(run);
    }))(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExField, InsertFieldWithFieldBuilderException)
{
    s_instance->InsertFieldWithFieldBuilderException();
}

} // namespace gtest_test

void ExField::PreserveIncludePicture(bool preserveIncludePictureField)
{
    //ExStart
    //ExFor:Field.Update(bool)
    //ExFor:LoadOptions.PreserveIncludePictureField
    //ExSummary:Shows how to preserve or discard INCLUDEPICTURE fields when loading a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto includePicture = System::ExplicitCast<Aspose::Words::Fields::FieldIncludePicture>(builder->InsertField(FieldType::FieldIncludePicture, true));
    includePicture->set_SourceFullName(get_ImageDir() + u"Transparent background logo.png");
    includePicture->Update(true);
    
    {
        auto docStream = System::MakeObject<System::IO::MemoryStream>();
        doc->Save(docStream, System::MakeObject<Aspose::Words::Saving::OoxmlSaveOptions>(SaveFormat::Docx));
        
        // We can set a flag in a LoadOptions object to decide whether to convert all INCLUDEPICTURE fields
        // into image shapes when loading a document that contains them.
        auto loadOptions = System::MakeObject<Aspose::Words::Loading::LoadOptions>();
        loadOptions->set_PreserveIncludePictureField(preserveIncludePictureField);
        
        doc = System::MakeObject<Aspose::Words::Document>(docStream, loadOptions);
        
        if (preserveIncludePictureField)
        {
            ASSERT_TRUE(doc->get_Range()->get_Fields()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
            {
                return f->get_Type() == FieldType::FieldIncludePicture;
            }))));
            
            doc->UpdateFields();
            doc->Save(get_ArtifactsDir() + u"Field.PreserveIncludePicture.docx");
        }
        else
        {
            ASSERT_FALSE(doc->get_Range()->get_Fields()->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
            {
                return f->get_Type() == FieldType::FieldIncludePicture;
            }))));
        }
    }
    //ExEnd
}

namespace gtest_test
{

using ExField_PreserveIncludePicture_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExField::PreserveIncludePicture)>::type;

struct ExField_PreserveIncludePicture : public ExField, public Aspose::Words::ApiExamples::ExField, public ::testing::WithParamInterface<ExField_PreserveIncludePicture_Args>
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

TEST_P(ExField_PreserveIncludePicture, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PreserveIncludePicture(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_PreserveIncludePicture, ::testing::ValuesIn(ExField_PreserveIncludePicture::TestCases()));

} // namespace gtest_test

void ExField::FieldFormat()
{
    //ExStart
    //ExFor:Field.Format
    //ExFor:Field.Update
    //ExFor:FieldFormat
    //ExFor:FieldFormat.DateTimeFormat
    //ExFor:FieldFormat.NumericFormat
    //ExFor:FieldFormat.GeneralFormats
    //ExFor:GeneralFormat
    //ExFor:GeneralFormatCollection
    //ExFor:GeneralFormatCollection.Add(GeneralFormat)
    //ExFor:GeneralFormatCollection.Count
    //ExFor:GeneralFormatCollection.Item(Int32)
    //ExFor:GeneralFormatCollection.Remove(GeneralFormat)
    //ExFor:GeneralFormatCollection.RemoveAt(Int32)
    //ExFor:GeneralFormatCollection.GetEnumerator
    //ExSummary:Shows how to format field results.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Use a document builder to insert a field that displays a result with no format applied.
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->InsertField(u"= 2 + 3");
    
    ASSERT_EQ(field->GetFieldCode(), u"= 2 + 3");
    ASSERT_EQ(field->get_Result(), u"5");
    
    // We can apply a format to a field's result using the field's properties.
    // Below are three types of formats that we can apply to a field's result.
    // 1 -  Numeric format:
    System::SharedPtr<Aspose::Words::Fields::FieldFormat> format = field->get_Format();
    format->set_NumericFormat(u"$###.00");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u"= 2 + 3 \\# $###.00");
    ASSERT_EQ(field->get_Result(), u"$  5.00");
    
    // 2 -  Date/time format:
    field = builder->InsertField(u"DATE");
    format = field->get_Format();
    format->set_DateTimeFormat(u"dddd, MMMM dd, yyyy");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), (u"DATE \\@ \"dddd, MMMM dd, yyyy\""));
    std::cout << System::String::Format(u"Today's date, in {0} format:\n\t{1}", format->get_DateTimeFormat(), field->get_Result()) << std::endl;
    
    // 3 -  General format:
    field = builder->InsertField(u"= 25 + 33");
    format = field->get_Format();
    format->get_GeneralFormats()->Add(GeneralFormat::LowercaseRoman);
    format->get_GeneralFormats()->Add(GeneralFormat::Upper);
    field->Update();
    
    int32_t index = 0;
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<Aspose::Words::Fields::GeneralFormat>> generalFormatEnumerator = format->get_GeneralFormats()->GetEnumerator();
        while (generalFormatEnumerator->MoveNext())
        {
            std::cout << System::String::Format(u"General format index {0}: {1}", index++, generalFormatEnumerator->get_Current()) << std::endl;
        }
    }
    
    ASSERT_EQ(field->GetFieldCode(), u"= 25 + 33 \\* roman \\* Upper");
    ASSERT_EQ(field->get_Result(), u"LVIII");
    ASSERT_EQ(format->get_GeneralFormats()->get_Count(), 2);
    ASSERT_EQ(format->get_GeneralFormats()->idx_get(0), GeneralFormat::LowercaseRoman);
    
    // We can remove our formats to revert the field's result to its original form.
    format->get_GeneralFormats()->Remove(GeneralFormat::LowercaseRoman);
    format->get_GeneralFormats()->RemoveAt(0);
    ASSERT_EQ(format->get_GeneralFormats()->get_Count(), 0);
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u"= 25 + 33  ");
    ASSERT_EQ(field->get_Result(), u"58");
    ASSERT_EQ(format->get_GeneralFormats()->get_Count(), 0);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, FieldFormat)
{
    s_instance->FieldFormat();
}

} // namespace gtest_test

void ExField::Unlink()
{
    //ExStart
    //ExFor:Document.UnlinkFields
    //ExSummary:Shows how to unlink all fields in the document.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Linked fields.docx"));
    
    doc->UnlinkFields();
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    System::String paraWithFields = DocumentHelper::GetParagraphText(doc, 0);
    
    ASSERT_EQ(paraWithFields, u"Fields.Docx   Элементы указателя не найдены.     1.\r");
}

namespace gtest_test
{

TEST_F(ExField, Unlink)
{
    s_instance->Unlink();
}

} // namespace gtest_test

void ExField::UnlinkAllFieldsInRange()
{
    //ExStart
    //ExFor:Range.UnlinkFields
    //ExSummary:Shows how to unlink all fields in a range.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Linked fields.docx"));
    
    auto newSection = System::ExplicitCast<Aspose::Words::Section>(System::ExplicitCast<Aspose::Words::Node>(doc->get_Sections()->idx_get(0))->Clone(true));
    doc->get_Sections()->Add(newSection);
    
    doc->get_Sections()->idx_get(1)->get_Range()->UnlinkFields();
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    System::String secWithFields = DocumentHelper::GetSectionText(doc, 1);
    
    ASSERT_TRUE(secWithFields.Trim().EndsWith(u"Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4."));
}

namespace gtest_test
{

TEST_F(ExField, UnlinkAllFieldsInRange)
{
    s_instance->UnlinkAllFieldsInRange();
}

} // namespace gtest_test

void ExField::UnlinkSingleField()
{
    //ExStart
    //ExFor:Field.Unlink
    //ExSummary:Shows how to unlink a field.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Linked fields.docx"));
    doc->get_Range()->get_Fields()->idx_get(1)->Unlink();
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    System::String paraWithFields = DocumentHelper::GetParagraphText(doc, 0);
    
    ASSERT_TRUE(paraWithFields.Trim().EndsWith(u"FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015"));
}

namespace gtest_test
{

TEST_F(ExField, UnlinkSingleField)
{
    s_instance->UnlinkSingleField();
}

} // namespace gtest_test

void ExField::UpdateTocPageNumbers()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Field sample - TOC.docx"));
    
    System::SharedPtr<Aspose::Words::Node> startNode = DocumentHelper::GetParagraph(doc, 2);
    System::SharedPtr<Aspose::Words::Node> endNode;
    
    System::SharedPtr<Aspose::Words::NodeCollection> paragraphCollection = doc->GetChildNodes(NodeType::Paragraph, true);
    
    for (auto&& para : System::IterateOver(paragraphCollection->LINQ_OfType<System::SharedPtr<Aspose::Words::Paragraph>>()))
    {
        for (auto&& run : System::IterateOver(para->get_Runs()->LINQ_OfType<System::SharedPtr<Aspose::Words::Run>>()))
        {
            if (run->get_Text().Contains(ControlChar::PageBreak()))
            {
                endNode = run;
                break;
            }
        }
    }
    
    if (startNode != nullptr && endNode != nullptr)
    {
        RemoveSequence(startNode, endNode);
        
        startNode->Remove();
        endNode->Remove();
    }
    
    System::SharedPtr<Aspose::Words::NodeCollection> fStart = doc->GetChildNodes(NodeType::FieldStart, true);
    
    for (auto&& field : System::IterateOver(fStart->LINQ_OfType<System::SharedPtr<Aspose::Words::Fields::FieldStart>>()))
    {
        Aspose::Words::Fields::FieldType fType = field->get_FieldType();
        if (fType == FieldType::FieldTOC)
        {
            auto para = System::ExplicitCast<Aspose::Words::Paragraph>(field->GetAncestor(NodeType::Paragraph));
            para->get_Range()->UpdateFields();
            break;
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"Field.UpdateTocPageNumbers.docx");
}

namespace gtest_test
{

TEST_F(ExField, UpdateTocPageNumbers)
{
    s_instance->UpdateTocPageNumbers();
}

} // namespace gtest_test

void ExField::FieldAdvance()
{
    //ExStart
    //ExFor:FieldAdvance
    //ExFor:FieldAdvance.DownOffset
    //ExFor:FieldAdvance.HorizontalPosition
    //ExFor:FieldAdvance.LeftOffset
    //ExFor:FieldAdvance.RightOffset
    //ExFor:FieldAdvance.UpOffset
    //ExFor:FieldAdvance.VerticalPosition
    //ExSummary:Shows how to insert an ADVANCE field, and edit its properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"This text is in its normal place.");
    
    // Below are two ways of using the ADVANCE field to adjust the position of text that follows it.
    // The effects of an ADVANCE field continue to be applied until the paragraph ends,
    // or another ADVANCE field updates the offset/coordinate values.
    // 1 -  Specify a directional offset:
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));
    ASSERT_EQ(field->get_Type(), FieldType::FieldAdvance);
    //ExSkip
    ASSERT_EQ(field->GetFieldCode(), u" ADVANCE ");
    //ExSkip
    field->set_RightOffset(u"5");
    field->set_UpOffset(u"5");
    
    ASSERT_EQ(field->GetFieldCode(), u" ADVANCE  \\r 5 \\u 5");
    
    builder->Write(u"This text will be moved up and to the right.");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));
    field->set_DownOffset(u"5");
    field->set_LeftOffset(u"100");
    
    ASSERT_EQ(field->GetFieldCode(), u" ADVANCE  \\d 5 \\l 100");
    
    builder->Writeln(u"This text is moved down and to the left, overlapping the previous text.");
    
    // 2 -  Move text to a position specified by coordinates:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAdvance>(builder->InsertField(FieldType::FieldAdvance, true));
    field->set_HorizontalPosition(u"-100");
    field->set_VerticalPosition(u"200");
    
    ASSERT_EQ(field->GetFieldCode(), u" ADVANCE  \\x -100 \\y 200");
    
    builder->Write(u"This text is in a custom position.");
    
    doc->Save(get_ArtifactsDir() + u"Field.ADVANCE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.ADVANCE.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAdvance>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldAdvance, u" ADVANCE  \\r 5 \\u 5", System::String::Empty, field);
    ASSERT_EQ(field->get_RightOffset(), u"5");
    ASSERT_EQ(field->get_UpOffset(), u"5");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAdvance>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldAdvance, u" ADVANCE  \\d 5 \\l 100", System::String::Empty, field);
    ASSERT_EQ(field->get_DownOffset(), u"5");
    ASSERT_EQ(field->get_LeftOffset(), u"100");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAdvance>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldAdvance, u" ADVANCE  \\x -100 \\y 200", System::String::Empty, field);
    ASSERT_EQ(field->get_HorizontalPosition(), u"-100");
    ASSERT_EQ(field->get_VerticalPosition(), u"200");
}

namespace gtest_test
{

TEST_F(ExField, FieldAdvance)
{
    s_instance->FieldAdvance();
}

} // namespace gtest_test

void ExField::FieldAddressBlock()
{
    //ExStart
    //ExFor:FieldAddressBlock.ExcludedCountryOrRegionName
    //ExFor:FieldAddressBlock.FormatAddressOnCountryOrRegion
    //ExFor:FieldAddressBlock.IncludeCountryOrRegionName
    //ExFor:FieldAddressBlock.LanguageId
    //ExFor:FieldAddressBlock.NameAndAddressFormat
    //ExSummary:Shows how to insert an ADDRESSBLOCK field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAddressBlock>(builder->InsertField(FieldType::FieldAddressBlock, true));
    
    ASSERT_EQ(field->GetFieldCode(), u" ADDRESSBLOCK ");
    
    // Setting this to "2" will include all countries and regions,
    // unless it is the one specified in the ExcludedCountryOrRegionName property.
    field->set_IncludeCountryOrRegionName(u"2");
    field->set_FormatAddressOnCountryOrRegion(true);
    field->set_ExcludedCountryOrRegionName(u"United States");
    field->set_NameAndAddressFormat(u"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");
    
    // By default, this property will contain the language ID of the first character of the document.
    // We can set a different culture for the field to format the result with like this.
    field->set_LanguageId(System::Convert::ToString(System::MakeObject<System::Globalization::CultureInfo>(u"en-US")->get_LCID()));
    
    ASSERT_EQ(field->GetFieldCode(), u" ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033");
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAddressBlock>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldAddressBlock, u" ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033", u"«AddressBlock»", field);
    ASSERT_EQ(field->get_IncludeCountryOrRegionName(), u"2");
    ASPOSE_ASSERT_EQ(field->get_FormatAddressOnCountryOrRegion(), true);
    ASSERT_EQ(field->get_ExcludedCountryOrRegionName(), u"United States");
    ASSERT_EQ(field->get_NameAndAddressFormat(), u"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");
    ASSERT_EQ(field->get_LanguageId(), u"1033");
}

namespace gtest_test
{

TEST_F(ExField, FieldAddressBlock)
{
    s_instance->FieldAddressBlock();
}

} // namespace gtest_test

void ExField::FieldCollection()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertField(u" DATE \\@ \"dddd, d MMMM yyyy\" ");
    builder->InsertField(u" TIME ");
    builder->InsertField(u" REVNUM ");
    builder->InsertField(u" AUTHOR  \"John Doe\" ");
    builder->InsertField(u" SUBJECT \"My Subject\" ");
    builder->InsertField(u" QUOTE \"Hello world!\" ");
    doc->UpdateFields();
    
    System::SharedPtr<Aspose::Words::Fields::FieldCollection> fields = doc->get_Range()->get_Fields();
    
    ASSERT_EQ(fields->get_Count(), 6);
    
    // Iterate over the field collection, and print contents and type
    // of every field using a custom visitor implementation.
    auto fieldVisitor = System::MakeObject<Aspose::Words::ApiExamples::ExField::FieldVisitor>();
    
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Fields::Field>>> fieldEnumerator = fields->GetEnumerator();
        while (fieldEnumerator->MoveNext())
        {
            if (fieldEnumerator->get_Current() != nullptr)
            {
                fieldEnumerator->get_Current()->get_Start()->Accept(fieldVisitor);
                System::SafeInvoke(fieldEnumerator->get_Current()->get_Separator(), [&](auto&& expr){return expr->Accept(fieldVisitor);});
                fieldEnumerator->get_Current()->get_End()->Accept(fieldVisitor);
            }
            else
            {
                std::cout << "There are no fields in the document." << std::endl;
            }
        }
    }
    
    std::cout << fieldVisitor->GetText() << std::endl;
    TestFieldCollection(fieldVisitor->GetText());
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldCollection)
{
    s_instance->FieldCollection();
}

} // namespace gtest_test

void ExField::RemoveFields()
{
    //ExStart
    //ExFor:FieldCollection
    //ExFor:FieldCollection.Count
    //ExFor:FieldCollection.Clear
    //ExFor:FieldCollection.Item(Int32)
    //ExFor:FieldCollection.Remove(Field)
    //ExFor:FieldCollection.RemoveAt(Int32)
    //ExFor:Field.Remove
    //ExSummary:Shows how to remove fields from a field collection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertField(u" DATE \\@ \"dddd, d MMMM yyyy\" ");
    builder->InsertField(u" TIME ");
    builder->InsertField(u" REVNUM ");
    builder->InsertField(u" AUTHOR  \"John Doe\" ");
    builder->InsertField(u" SUBJECT \"My Subject\" ");
    builder->InsertField(u" QUOTE \"Hello world!\" ");
    doc->UpdateFields();
    
    System::SharedPtr<Aspose::Words::Fields::FieldCollection> fields = doc->get_Range()->get_Fields();
    
    ASSERT_EQ(fields->get_Count(), 6);
    
    // Below are four ways of removing fields from a field collection.
    // 1 -  Get a field to remove itself:
    fields->idx_get(0)->Remove();
    ASSERT_EQ(fields->get_Count(), 5);
    
    // 2 -  Get the collection to remove a field that we pass to its removal method:
    System::SharedPtr<Aspose::Words::Fields::Field> lastField = fields->idx_get(3);
    fields->Remove(lastField);
    ASSERT_EQ(fields->get_Count(), 4);
    
    // 3 -  Remove a field from a collection at an index:
    fields->RemoveAt(2);
    ASSERT_EQ(fields->get_Count(), 3);
    
    // 4 -  Remove all the fields from the collection at once:
    fields->Clear();
    ASSERT_EQ(fields->get_Count(), 0);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, RemoveFields)
{
    s_instance->RemoveFields();
}

} // namespace gtest_test

void ExField::FieldCompare()
{
    //ExStart
    //ExFor:FieldCompare
    //ExFor:FieldCompare.ComparisonOperator
    //ExFor:FieldCompare.LeftExpression
    //ExFor:FieldCompare.RightExpression
    //ExSummary:Shows how to compare expressions using a COMPARE field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldCompare>(builder->InsertField(FieldType::FieldCompare, true));
    field->set_LeftExpression(u"3");
    field->set_ComparisonOperator(u"<");
    field->set_RightExpression(u"2");
    field->Update();
    
    // The COMPARE field displays a "0" or a "1", depending on its statement's truth.
    // The result of this statement is false so that this field will display a "0".
    ASSERT_EQ(field->GetFieldCode(), u" COMPARE  3 < 2");
    ASSERT_EQ(field->get_Result(), u"0");
    
    builder->Writeln();
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldCompare>(builder->InsertField(FieldType::FieldCompare, true));
    field->set_LeftExpression(u"5");
    field->set_ComparisonOperator(u"=");
    field->set_RightExpression(u"2 + 3");
    field->Update();
    
    // This field displays a "1" since the statement is true.
    ASSERT_EQ(field->GetFieldCode(), u" COMPARE  5 = \"2 + 3\"");
    ASSERT_EQ(field->get_Result(), u"1");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.COMPARE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.COMPARE.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldCompare>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldCompare, u" COMPARE  3 < 2", u"0", field);
    ASSERT_EQ(field->get_LeftExpression(), u"3");
    ASSERT_EQ(field->get_ComparisonOperator(), u"<");
    ASSERT_EQ(field->get_RightExpression(), u"2");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldCompare>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldCompare, u" COMPARE  5 = \"2 + 3\"", u"1", field);
    ASSERT_EQ(field->get_LeftExpression(), u"5");
    ASSERT_EQ(field->get_ComparisonOperator(), u"=");
    ASSERT_EQ(field->get_RightExpression(), u"\"2 + 3\"");
}

namespace gtest_test
{

TEST_F(ExField, FieldCompare)
{
    s_instance->FieldCompare();
}

} // namespace gtest_test

void ExField::FieldIf()
{
    //ExStart
    //ExFor:FieldIf
    //ExFor:FieldIf.ComparisonOperator
    //ExFor:FieldIf.EvaluateCondition
    //ExFor:FieldIf.FalseText
    //ExFor:FieldIf.LeftExpression
    //ExFor:FieldIf.RightExpression
    //ExFor:FieldIf.TrueText
    //ExFor:FieldIfComparisonResult
    //ExSummary:Shows how to insert an IF field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Statement 1: ");
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldIf>(builder->InsertField(FieldType::FieldIf, true));
    field->set_LeftExpression(u"0");
    field->set_ComparisonOperator(u"=");
    field->set_RightExpression(u"1");
    
    // The IF field will display a string from either its "TrueText" property,
    // or its "FalseText" property, depending on the truth of the statement that we have constructed.
    field->set_TrueText(u"True");
    field->set_FalseText(u"False");
    field->Update();
    
    // In this case, "0 = 1" is incorrect, so the displayed result will be "False".
    ASSERT_EQ(field->GetFieldCode(), u" IF  0 = 1 True False");
    ASSERT_EQ(field->EvaluateCondition(), FieldIfComparisonResult::False);
    ASSERT_EQ(field->get_Result(), u"False");
    
    builder->Write(u"\nStatement 2: ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldIf>(builder->InsertField(FieldType::FieldIf, true));
    field->set_LeftExpression(u"5");
    field->set_ComparisonOperator(u"=");
    field->set_RightExpression(u"2 + 3");
    field->set_TrueText(u"True");
    field->set_FalseText(u"False");
    field->Update();
    
    // This time the statement is correct, so the displayed result will be "True".
    ASSERT_EQ(field->GetFieldCode(), u" IF  5 = \"2 + 3\" True False");
    ASSERT_EQ(field->EvaluateCondition(), FieldIfComparisonResult::True);
    ASSERT_EQ(field->get_Result(), u"True");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.IF.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.IF.docx"));
    field = System::ExplicitCast<Aspose::Words::Fields::FieldIf>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldIf, u" IF  0 = 1 True False", u"False", field);
    ASSERT_EQ(field->get_LeftExpression(), u"0");
    ASSERT_EQ(field->get_ComparisonOperator(), u"=");
    ASSERT_EQ(field->get_RightExpression(), u"1");
    ASSERT_EQ(field->get_TrueText(), u"True");
    ASSERT_EQ(field->get_FalseText(), u"False");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldIf>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldIf, u" IF  5 = \"2 + 3\" True False", u"True", field);
    ASSERT_EQ(field->get_LeftExpression(), u"5");
    ASSERT_EQ(field->get_ComparisonOperator(), u"=");
    ASSERT_EQ(field->get_RightExpression(), u"\"2 + 3\"");
    ASSERT_EQ(field->get_TrueText(), u"True");
    ASSERT_EQ(field->get_FalseText(), u"False");
}

namespace gtest_test
{

TEST_F(ExField, FieldIf)
{
    s_instance->FieldIf();
}

} // namespace gtest_test

void ExField::FieldAutoNum()
{
    //ExStart
    //ExFor:FieldAutoNum
    //ExFor:FieldAutoNum.SeparatorCharacter
    //ExSummary:Shows how to number paragraphs using autonum fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Each AUTONUM field displays the current value of a running count of AUTONUM fields,
    // allowing us to automatically number items like a numbered list.
    // This field will display a number "1.".
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAutoNum>(builder->InsertField(FieldType::FieldAutoNum, true));
    builder->Writeln(u"\tParagraph 1.");
    
    ASSERT_EQ(field->GetFieldCode(), u" AUTONUM ");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAutoNum>(builder->InsertField(FieldType::FieldAutoNum, true));
    builder->Writeln(u"\tParagraph 2.");
    
    // The separator character, which appears in the field result immediately after the number,is a full stop by default.
    // If we leave this property null, our second AUTONUM field will display "2." in the document.
    ASSERT_TRUE(System::TestTools::IsNull(field->get_SeparatorCharacter()));
    
    // We can set this property to apply the first character of its string as the new separator character.
    // In this case, our AUTONUM field will now display "2:".
    field->set_SeparatorCharacter(u":");
    
    ASSERT_EQ(field->GetFieldCode(), u" AUTONUM  \\s :");
    
    doc->Save(get_ArtifactsDir() + u"Field.AUTONUM.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.AUTONUM.docx"));
    
    TestUtil::VerifyField(FieldType::FieldAutoNum, u" AUTONUM ", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
    TestUtil::VerifyField(FieldType::FieldAutoNum, u" AUTONUM  \\s :", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(1));
}

namespace gtest_test
{

TEST_F(ExField, FieldAutoNum)
{
    s_instance->FieldAutoNum();
}

} // namespace gtest_test

void ExField::FieldAutoNumLgl()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    const System::String fillerText = System::String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") + u"\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";
    
    // AUTONUMLGL fields display a number that increments at each AUTONUMLGL field within its current heading level.
    // These fields maintain a separate count for each heading level,
    // and each field also displays the AUTONUMLGL field counts for all heading levels below its own.
    // Changing the count for any heading level resets the counts for all levels above that level to 1.
    // This allows us to organize our document in the form of an outline list.
    // This is the first AUTONUMLGL field at a heading level of 1, displaying "1." in the document.
    InsertNumberedClause(builder, u"\tHeading 1", fillerText, StyleIdentifier::Heading1);
    
    // This is the second AUTONUMLGL field at a heading level of 1, so it will display "2.".
    InsertNumberedClause(builder, u"\tHeading 2", fillerText, StyleIdentifier::Heading1);
    
    // This is the first AUTONUMLGL field at a heading level of 2,
    // and the AUTONUMLGL count for the heading level below it is "2", so it will display "2.1.".
    InsertNumberedClause(builder, u"\tHeading 3", fillerText, StyleIdentifier::Heading2);
    
    // This is the first AUTONUMLGL field at a heading level of 3.
    // Working in the same way as the field above, it will display "2.1.1.".
    InsertNumberedClause(builder, u"\tHeading 4", fillerText, StyleIdentifier::Heading3);
    
    // This field is at a heading level of 2, and its respective AUTONUMLGL count is at 2, so the field will display "2.2.".
    InsertNumberedClause(builder, u"\tHeading 5", fillerText, StyleIdentifier::Heading2);
    
    // Incrementing the AUTONUMLGL count for a heading level below this one
    // has reset the count for this level so that this field will display "2.2.1.".
    InsertNumberedClause(builder, u"\tHeading 6", fillerText, StyleIdentifier::Heading3);
    
    for (auto&& field : System::IterateOver<Aspose::Words::Fields::FieldAutoNumLgl>(doc->get_Range()->get_Fields()->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
    {
        return f->get_Type() == FieldType::FieldAutoNumLegal;
    })))->LINQ_ToList()))
    {
        // The separator character, which appears in the field result immediately after the number,
        // is a full stop by default. If we leave this property null,
        // our last AUTONUMLGL field will display "2.2.1." in the document.
        ASSERT_TRUE(System::TestTools::IsNull(field->get_SeparatorCharacter()));
        
        // Setting a custom separator character and removing the trailing period
        // will change that field's appearance from "2.2.1." to "2:2:1".
        // We will apply this to all the fields that we have created.
        field->set_SeparatorCharacter(u":");
        field->set_RemoveTrailingPeriod(true);
        ASSERT_EQ(field->GetFieldCode(), u" AUTONUMLGL  \\s : \\e");
    }
    
    doc->Save(get_ArtifactsDir() + u"Field.AUTONUMLGL.docx");
    TestFieldAutoNumLgl(doc);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldAutoNumLgl)
{
    s_instance->FieldAutoNumLgl();
}

} // namespace gtest_test

void ExField::FieldAutoNumOut()
{
    //ExStart
    //ExFor:FieldAutoNumOut
    //ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // AUTONUMOUT fields display a number that increments at each AUTONUMOUT field.
    // Unlike AUTONUM fields, AUTONUMOUT fields use the outline numbering scheme,
    // which we can define in Microsoft Word via Format -> Bullets & Numbering -> "Outline Numbered".
    // This allows us to automatically number items like a numbered list.
    // LISTNUM fields are a newer alternative to AUTONUMOUT fields.
    // This field will display "1.".
    builder->InsertField(FieldType::FieldAutoNumOutline, true);
    builder->Writeln(u"\tParagraph 1.");
    
    // This field will display "2.".
    builder->InsertField(FieldType::FieldAutoNumOutline, true);
    builder->Writeln(u"\tParagraph 2.");
    
    for (auto&& field : System::IterateOver<Aspose::Words::Fields::FieldAutoNumOut>(doc->get_Range()->get_Fields()->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
    {
        return f->get_Type() == FieldType::FieldAutoNumOutline;
    })))->LINQ_ToList()))
    {
        ASSERT_EQ(field->GetFieldCode(), u" AUTONUMOUT ");
    }
    
    doc->Save(get_ArtifactsDir() + u"Field.AUTONUMOUT.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.AUTONUMOUT.docx"));
    
    for (auto&& field : System::IterateOver(doc->get_Range()->get_Fields()))
    {
        TestUtil::VerifyField(FieldType::FieldAutoNumOutline, u" AUTONUMOUT ", System::String::Empty, field);
    }
}

namespace gtest_test
{

TEST_F(ExField, FieldAutoNumOut)
{
    s_instance->FieldAutoNumOut();
}

} // namespace gtest_test

void ExField::FieldAutoText()
{
    //ExStart
    //ExFor:FieldAutoText
    //ExFor:FieldAutoText.EntryName
    //ExFor:FieldOptions.BuiltInTemplatesPaths
    //ExFor:FieldGlossary
    //ExFor:FieldGlossary.EntryName
    //ExSummary:Shows how to display a building block with AUTOTEXT and GLOSSARY fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a glossary document and add an AutoText building block to it.
    doc->set_GlossaryDocument(System::MakeObject<Aspose::Words::BuildingBlocks::GlossaryDocument>());
    auto buildingBlock = System::MakeObject<Aspose::Words::BuildingBlocks::BuildingBlock>(doc->get_GlossaryDocument());
    buildingBlock->set_Name(u"MyBlock");
    buildingBlock->set_Gallery(BuildingBlockGallery::AutoText);
    buildingBlock->set_Category(u"General");
    buildingBlock->set_Description(u"MyBlock description");
    buildingBlock->set_Behavior(BuildingBlockBehavior::Paragraph);
    doc->get_GlossaryDocument()->AppendChild<System::SharedPtr<Aspose::Words::BuildingBlocks::BuildingBlock>>(buildingBlock);
    
    // Create a source and add it as text to our building block.
    auto buildingBlockSource = System::MakeObject<Aspose::Words::Document>();
    auto buildingBlockSourceBuilder = System::MakeObject<Aspose::Words::DocumentBuilder>(buildingBlockSource);
    buildingBlockSourceBuilder->Writeln(u"Hello World!");
    
    System::SharedPtr<Aspose::Words::Node> buildingBlockContent = doc->get_GlossaryDocument()->ImportNode(buildingBlockSource->get_FirstSection(), true);
    buildingBlock->AppendChild<System::SharedPtr<Aspose::Words::Node>>(buildingBlockContent);
    
    // Set a file which contains parts that our document, or its attached template may not contain.
    doc->get_FieldOptions()->set_BuiltInTemplatesPaths(System::MakeArray<System::String>({
        get_MyDir() + u"Busniess brochure.dotx"}));
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two ways to use fields to display the contents of our building block.
    // 1 -  Using an AUTOTEXT field:
    auto fieldAutoText = System::ExplicitCast<Aspose::Words::Fields::FieldAutoText>(builder->InsertField(FieldType::FieldAutoText, true));
    fieldAutoText->set_EntryName(u"MyBlock");
    
    ASSERT_EQ(fieldAutoText->GetFieldCode(), u" AUTOTEXT  MyBlock");
    
    // 2 -  Using a GLOSSARY field:
    auto fieldGlossary = System::ExplicitCast<Aspose::Words::Fields::FieldGlossary>(builder->InsertField(FieldType::FieldGlossary, true));
    fieldGlossary->set_EntryName(u"MyBlock");
    
    ASSERT_EQ(fieldGlossary->GetFieldCode(), u" GLOSSARY  MyBlock");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.AUTOTEXT.GLOSSARY.dotx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.AUTOTEXT.GLOSSARY.dotx"));
    
    ASSERT_EQ(doc->get_FieldOptions()->get_BuiltInTemplatesPaths()->get_Length(), 0);
    
    fieldAutoText = System::ExplicitCast<Aspose::Words::Fields::FieldAutoText>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldAutoText, u" AUTOTEXT  MyBlock", u"Hello World!\r", fieldAutoText);
    ASSERT_EQ(fieldAutoText->get_EntryName(), u"MyBlock");
    
    fieldGlossary = System::ExplicitCast<Aspose::Words::Fields::FieldGlossary>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldGlossary, u" GLOSSARY  MyBlock", u"Hello World!\r", fieldGlossary);
    ASSERT_EQ(fieldGlossary->get_EntryName(), u"MyBlock");
}

namespace gtest_test
{

TEST_F(ExField, FieldAutoText)
{
    s_instance->FieldAutoText();
}

} // namespace gtest_test

void ExField::FieldAutoTextList()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a glossary document and populate it with auto text entries.
    doc->set_GlossaryDocument(System::MakeObject<Aspose::Words::BuildingBlocks::GlossaryDocument>());
    AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 1", u"Contents of AutoText 1");
    AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 2", u"Contents of AutoText 2");
    AppendAutoTextEntry(doc->get_GlossaryDocument(), u"AutoText 3", u"Contents of AutoText 3");
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an AUTOTEXTLIST field and set the text that the field will display in Microsoft Word.
    // Set the text to prompt the user to right-click this field to select an AutoText building block,
    // whose contents the field will display.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAutoTextList>(builder->InsertField(FieldType::FieldAutoTextList, true));
    field->set_EntryName(u"Right click here to select an AutoText block");
    field->set_ListStyle(u"Heading 1");
    field->set_ScreenTip(u"Hover tip text for AutoTextList goes here");
    
    ASSERT_EQ(field->GetFieldCode(), System::String(u" AUTOTEXTLIST  \"Right click here to select an AutoText block\" ") + u"\\s \"Heading 1\" " + u"\\t \"Hover tip text for AutoTextList goes here\"");
    
    doc->Save(get_ArtifactsDir() + u"Field.AUTOTEXTLIST.dotx");
    TestFieldAutoTextList(doc);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldAutoTextList)
{
    s_instance->FieldAutoTextList();
}

} // namespace gtest_test

void ExField::FieldListNum()
{
    //ExStart
    //ExFor:FieldListNum
    //ExFor:FieldListNum.HasListName
    //ExFor:FieldListNum.ListLevel
    //ExFor:FieldListNum.ListName
    //ExFor:FieldListNum.StartingNumber
    //ExSummary:Shows how to number paragraphs with LISTNUM fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // LISTNUM fields display a number that increments at each LISTNUM field.
    // These fields also have a variety of options that allow us to use them to emulate numbered lists.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
    
    // Lists start counting at 1 by default, but we can set this number to a different value, such as 0.
    // This field will display "0)".
    field->set_StartingNumber(u"0");
    builder->Writeln(u"Paragraph 1");
    
    ASSERT_EQ(field->GetFieldCode(), u" LISTNUM  \\s 0");
    
    // LISTNUM fields maintain separate counts for each list level.
    // Inserting a LISTNUM field in the same paragraph as another LISTNUM field
    // increases the list level instead of the count.
    // The next field will continue the count we started above and display a value of "1" at list level 1.
    builder->InsertField(FieldType::FieldListNum, true);
    
    // This field will start a count at list level 2. It will display a value of "1".
    builder->InsertField(FieldType::FieldListNum, true);
    
    // This field will start a count at list level 3. It will display a value of "1".
    // Different list levels have different formatting,
    // so these fields combined will display a value of "1)a)i)".
    builder->InsertField(FieldType::FieldListNum, true);
    builder->Writeln(u"Paragraph 2");
    
    // The next LISTNUM field that we insert will continue the count at the list level
    // that the previous LISTNUM field was on.
    // We can use the "ListLevel" property to jump to a different list level.
    // If this LISTNUM field stayed on list level 3, it would display "ii)",
    // but, since we have moved it to list level 2, it carries on the count at that level and displays "b)".
    field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
    field->set_ListLevel(u"2");
    builder->Writeln(u"Paragraph 3");
    
    ASSERT_EQ(field->GetFieldCode(), u" LISTNUM  \\l 2");
    
    // We can set the ListName property to get the field to emulate a different AUTONUM field type.
    // "NumberDefault" emulates AUTONUM, "OutlineDefault" emulates AUTONUMOUT,
    // and "LegalDefault" emulates AUTONUMLGL fields.
    // The "OutlineDefault" list name with 1 as the starting number will result in displaying "I.".
    field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
    field->set_StartingNumber(u"1");
    field->set_ListName(u"OutlineDefault");
    builder->Writeln(u"Paragraph 4");
    
    ASSERT_TRUE(field->get_HasListName());
    ASSERT_EQ(field->GetFieldCode(), u" LISTNUM  OutlineDefault \\s 1");
    
    // The ListName does not carry over from the previous field, so we will need to set it for each new field.
    // This field continues the count with the different list name and displays "II.".
    field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(builder->InsertField(FieldType::FieldListNum, true));
    field->set_ListName(u"OutlineDefault");
    builder->Writeln(u"Paragraph 5");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.LISTNUM.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.LISTNUM.docx"));
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 7);
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM  \\s 0", System::String::Empty, field);
    ASSERT_EQ(field->get_StartingNumber(), u"0");
    ASSERT_TRUE(System::TestTools::IsNull(field->get_ListLevel()));
    ASSERT_FALSE(field->get_HasListName());
    ASSERT_TRUE(System::TestTools::IsNull(field->get_ListName()));
    
    for (int32_t i = 1; i < 4; i++)
    {
        field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(doc->get_Range()->get_Fields()->idx_get(i));
        
        TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM ", System::String::Empty, field);
        ASSERT_TRUE(System::TestTools::IsNull(field->get_StartingNumber()));
        ASSERT_TRUE(System::TestTools::IsNull(field->get_ListLevel()));
        ASSERT_FALSE(field->get_HasListName());
        ASSERT_TRUE(System::TestTools::IsNull(field->get_ListName()));
    }
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM  \\l 2", System::String::Empty, field);
    ASSERT_TRUE(System::TestTools::IsNull(field->get_StartingNumber()));
    ASSERT_EQ(field->get_ListLevel(), u"2");
    ASSERT_FALSE(field->get_HasListName());
    ASSERT_TRUE(System::TestTools::IsNull(field->get_ListName()));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldListNum>(doc->get_Range()->get_Fields()->idx_get(5));
    
    TestUtil::VerifyField(FieldType::FieldListNum, u" LISTNUM  OutlineDefault \\s 1", System::String::Empty, field);
    ASSERT_EQ(field->get_StartingNumber(), u"1");
    ASSERT_TRUE(System::TestTools::IsNull(field->get_ListLevel()));
    ASSERT_TRUE(field->get_HasListName());
    ASSERT_EQ(field->get_ListName(), u"OutlineDefault");
}

namespace gtest_test
{

TEST_F(ExField, FieldListNum)
{
    s_instance->FieldListNum();
}

} // namespace gtest_test

void ExField::FieldToc()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartBookmark(u"MyBookmark");
    
    // Insert a TOC field, which will compile all headings into a table of contents.
    // For each heading, this field will create a line with the text in that heading style to the left,
    // and the page the heading appears on to the right.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
    
    // Use the BookmarkName property to only list headings
    // that appear within the bounds of a bookmark with the "MyBookmark" name.
    field->set_BookmarkName(u"MyBookmark");
    
    // Text with a built-in heading style, such as "Heading 1", applied to it will count as a heading.
    // We can name additional styles to be picked up as headings by the TOC in this property and their TOC levels.
    field->set_CustomStyles(u"Quote; 6; Intense Quote; 7");
    
    // By default, Styles/TOC levels are separated in the CustomStyles property by a comma,
    // but we can set a custom delimiter in this property.
    doc->get_FieldOptions()->set_CustomTocStyleSeparator(u";");
    
    // Configure the field to exclude any headings that have TOC levels outside of this range.
    field->set_HeadingLevelRange(u"1-3");
    
    // The TOC will not display the page numbers of headings whose TOC levels are within this range.
    field->set_PageNumberOmittingLevelRange(u"2-5");
    
    // Set a custom string that will separate every heading from its page number.
    field->set_EntrySeparator(u"-");
    field->set_InsertHyperlinks(true);
    field->set_HideInWebLayout(false);
    field->set_PreserveLineBreaks(true);
    field->set_PreserveTabs(true);
    field->set_UseParagraphOutlineLevel(false);
    
    InsertNewPageWithHeading(builder, u"First entry", u"Heading 1");
    builder->Writeln(u"Paragraph text.");
    InsertNewPageWithHeading(builder, u"Second entry", u"Heading 1");
    InsertNewPageWithHeading(builder, u"Third entry", u"Quote");
    InsertNewPageWithHeading(builder, u"Fourth entry", u"Intense Quote");
    
    // These two headings will have the page numbers omitted because they are within the "2-5" range.
    InsertNewPageWithHeading(builder, u"Fifth entry", u"Heading 2");
    InsertNewPageWithHeading(builder, u"Sixth entry", u"Heading 3");
    
    // This entry does not appear because "Heading 4" is outside of the "1-3" range that we have set earlier.
    InsertNewPageWithHeading(builder, u"Seventh entry", u"Heading 4");
    
    builder->EndBookmark(u"MyBookmark");
    builder->Writeln(u"Paragraph text.");
    
    // This entry does not appear because it is outside the bookmark specified by the TOC.
    InsertNewPageWithHeading(builder, u"Eighth entry", u"Heading 1");
    
    ASSERT_EQ(field->GetFieldCode(), u" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w");
    
    field->UpdatePageNumbers();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.TOC.docx");
    TestFieldToc(doc);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldToc)
{
    s_instance->FieldToc();
}

} // namespace gtest_test

void ExField::FieldTocEntryIdentifier()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a TOC field, which will compile all TC fields into a table of contents.
    auto fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
    
    // Configure the field only to pick up TC entries of the "A" type, and an entry-level between 1 and 3.
    fieldToc->set_EntryIdentifier(u"A");
    fieldToc->set_EntryLevelRange(u"1-3");
    
    ASSERT_EQ(fieldToc->GetFieldCode(), u" TOC  \\f A \\l 1-3");
    
    // These two entries will appear in the table.
    builder->InsertBreak(BreakType::PageBreak);
    InsertTocEntry(builder, u"TC field 1", u"A", u"1");
    InsertTocEntry(builder, u"TC field 2", u"A", u"2");
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(1)->GetFieldCode(), u" TC  \"TC field 1\" \\n \\f A \\l 1");
    
    // This entry will be omitted from the table because it has a different type from "A".
    InsertTocEntry(builder, u"TC field 3", u"B", u"1");
    
    // This entry will be omitted from the table because it has an entry-level outside of the 1-3 range.
    InsertTocEntry(builder, u"TC field 4", u"A", u"5");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.TC.docx");
    TestFieldTocEntryIdentifier(doc);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldTocEntryIdentifier)
{
    s_instance->FieldTocEntryIdentifier();
}

} // namespace gtest_test

void ExField::TocSeqPrefix()
{
    //ExStart
    //ExFor:FieldToc
    //ExFor:FieldToc.TableOfFiguresLabel
    //ExFor:FieldToc.PrefixedSequenceIdentifier
    //ExFor:FieldToc.SequenceSeparator
    //ExFor:FieldSeq
    //ExFor:FieldSeq.SequenceIdentifier
    //ExSummary:Shows how to populate a TOC field with entries using SEQ fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
    // Each entry contains the paragraph that includes the SEQ field and the page's number that the field appears on.
    auto fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
    
    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Use the "TableOfFiguresLabel" property to name a main sequence for the TOC.
    // Now, this TOC will only create entries out of SEQ fields with their "SequenceIdentifier" set to "MySequence".
    fieldToc->set_TableOfFiguresLabel(u"MySequence");
    
    // We can name another SEQ field sequence in the "PrefixedSequenceIdentifier" property.
    // SEQ fields from this prefix sequence will not create TOC entries.
    // Every TOC entry created from a main sequence SEQ field will now also display the count that
    // the prefix sequence is currently on at the primary sequence SEQ field that made the entry.
    fieldToc->set_PrefixedSequenceIdentifier(u"PrefixSequence");
    
    // Each TOC entry will display the prefix sequence count immediately to the left
    // of the page number that the main sequence SEQ field appears on.
    // We can specify a custom separator that will appear between these two numbers.
    fieldToc->set_SequenceSeparator(u">");
    
    ASSERT_EQ(fieldToc->GetFieldCode(), u" TOC  \\c MySequence \\s PrefixSequence \\d >");
    
    builder->InsertBreak(BreakType::PageBreak);
    
    // There are two ways of using SEQ fields to populate this TOC.
    // 1 -  Inserting a SEQ field that belongs to the TOC's prefix sequence:
    // This field will increment the SEQ sequence count for the "PrefixSequence" by 1.
    // Since this field does not belong to the main sequence identified
    // by the "TableOfFiguresLabel" property of the TOC, it will not appear as an entry.
    auto fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"PrefixSequence");
    builder->InsertParagraph();
    
    ASSERT_EQ(fieldSeq->GetFieldCode(), u" SEQ  PrefixSequence");
    
    // 2 -  Inserting a SEQ field that belongs to the TOC's main sequence:
    // This SEQ field will create an entry in the TOC.
    // The TOC entry will contain the paragraph that the SEQ field is in and the number of the page that it appears on.
    // This entry will also display the count that the prefix sequence is currently at,
    // separated from the page number by the value in the TOC's SeqenceSeparator property.
    // The "PrefixSequence" count is at 1, this main sequence SEQ field is on page 2,
    // and the separator is ">", so entry will display "1>2".
    builder->Write(u"First TOC entry, MySequence #");
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    
    ASSERT_EQ(fieldSeq->GetFieldCode(), u" SEQ  MySequence");
    
    // Insert a page, advance the prefix sequence by 2, and insert a SEQ field to create a TOC entry afterwards.
    // The prefix sequence is now at 2, and the main sequence SEQ field is on page 3,
    // so the TOC entry will display "2>3" at its page count.
    builder->InsertBreak(BreakType::PageBreak);
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"PrefixSequence");
    builder->InsertParagraph();
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    builder->Write(u"Second TOC entry, MySequence #");
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.TOC.SEQ.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.TOC.SEQ.docx"));
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 9);
    
    fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
    std::cout << fieldToc->get_DisplayResult() << std::endl;
    TestUtil::VerifyField(FieldType::FieldTOC, u" TOC  \\c MySequence \\s PrefixSequence \\d >", System::String(u"First TOC entry, MySequence #12\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r2") + u"Second TOC entry, MySequence #\t\u0013 SEQ PrefixSequence _Toc256000001 \\* ARABIC \u00142\u0015>\u0013 PAGEREF _Toc256000001 \\h \u00143\u0015\r", fieldToc);
    ASSERT_EQ(fieldToc->get_TableOfFiguresLabel(), u"MySequence");
    ASSERT_EQ(fieldToc->get_PrefixedSequenceIdentifier(), u"PrefixSequence");
    ASSERT_EQ(fieldToc->get_SequenceSeparator(), u">");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ PrefixSequence _Toc256000000 \\* ARABIC ", u"1", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"PrefixSequence");
    
    // Byproduct field created by Aspose.Words
    auto fieldPageRef = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF _Toc256000000 \\h ", u"2", fieldPageRef);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"PrefixSequence");
    ASSERT_EQ(fieldPageRef->get_BookmarkName(), u"_Toc256000000");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ PrefixSequence _Toc256000001 \\* ARABIC ", u"2", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"PrefixSequence");
    
    fieldPageRef = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF _Toc256000001 \\h ", u"3", fieldPageRef);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"PrefixSequence");
    ASSERT_EQ(fieldPageRef->get_BookmarkName(), u"_Toc256000001");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(5));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  PrefixSequence", u"1", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"PrefixSequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(6));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"1", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(7));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  PrefixSequence", u"2", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"PrefixSequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(8));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"2", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
}

namespace gtest_test
{

TEST_F(ExField, TocSeqPrefix)
{
    s_instance->TocSeqPrefix();
}

} // namespace gtest_test

void ExField::TocSeqNumbering()
{
    //ExStart
    //ExFor:FieldSeq
    //ExFor:FieldSeq.InsertNextNumber
    //ExFor:FieldSeq.ResetHeadingLevel
    //ExFor:FieldSeq.ResetNumber
    //ExFor:FieldSeq.SequenceIdentifier
    //ExSummary:Shows create numbering using SEQ fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Insert a SEQ field that will display the current count value of "MySequence",
    // after using the "ResetNumber" property to set it to 100.
    builder->Write(u"#");
    auto fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    fieldSeq->set_ResetNumber(u"100");
    fieldSeq->Update();
    
    ASSERT_EQ(fieldSeq->GetFieldCode(), u" SEQ  MySequence \\r 100");
    ASSERT_EQ(fieldSeq->get_Result(), u"100");
    
    // Display the next number in this sequence with another SEQ field.
    builder->Write(u", #");
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    fieldSeq->Update();
    
    ASSERT_EQ(fieldSeq->get_Result(), u"101");
    
    // Insert a level 1 heading.
    builder->InsertBreak(BreakType::ParagraphBreak);
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Heading 1"));
    builder->Writeln(u"This level 1 heading will reset MySequence to 1");
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
    
    // Insert another SEQ field from the same sequence and configure it to reset the count at every heading with 1.
    builder->Write(u"\n#");
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    fieldSeq->set_ResetHeadingLevel(u"1");
    fieldSeq->Update();
    
    // The above heading is a level 1 heading, so the count for this sequence is reset to 1.
    ASSERT_EQ(fieldSeq->GetFieldCode(), u" SEQ  MySequence \\s 1");
    ASSERT_EQ(fieldSeq->get_Result(), u"1");
    
    // Move to the next number of this sequence.
    builder->Write(u", #");
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    fieldSeq->set_InsertNextNumber(true);
    fieldSeq->Update();
    
    ASSERT_EQ(fieldSeq->GetFieldCode(), u" SEQ  MySequence \\n");
    ASSERT_EQ(fieldSeq->get_Result(), u"2");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.SEQ.ResetNumbering.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SEQ.ResetNumbering.docx"));
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 4);
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence \\r 100", u"100", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"101", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence \\s 1", u"1", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence \\n", u"2", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
}

namespace gtest_test
{

TEST_F(ExField, TocSeqNumbering)
{
    s_instance->TocSeqNumbering();
}

} // namespace gtest_test

void ExField::TocSeqBookmark()
{
    //ExStart
    //ExFor:FieldSeq
    //ExFor:FieldSeq.BookmarkName
    //ExSummary:Shows how to combine table of contents and sequence fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
    // Each entry contains the paragraph that contains the SEQ field,
    // and the number of the page that the field appears on.
    auto fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
    
    // Configure this TOC field to have a SequenceIdentifier property with a value of "MySequence".
    fieldToc->set_TableOfFiguresLabel(u"MySequence");
    
    // Configure this TOC field to only pick up SEQ fields that are within the bounds of a bookmark
    // named "TOCBookmark".
    fieldToc->set_BookmarkName(u"TOCBookmark");
    builder->InsertBreak(BreakType::PageBreak);
    
    ASSERT_EQ(fieldToc->GetFieldCode(), u" TOC  \\c MySequence \\b TOCBookmark");
    
    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Insert a SEQ field that has a sequence identifier that matches the TOC's
    // TableOfFiguresLabel property. This field will not create an entry in the TOC since it is outside
    // the bookmark's bounds designated by "BookmarkName".
    builder->Write(u"MySequence #");
    auto fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    builder->Writeln(u", will not show up in the TOC because it is outside of the bookmark.");
    
    builder->StartBookmark(u"TOCBookmark");
    
    // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bookmark's bounds.
    // The paragraph that contains this field will show up in the TOC as an entry.
    builder->Write(u"MySequence #");
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    builder->Writeln(u", will show up in the TOC next to the entry for the above caption.");
    
    // This SEQ field's sequence does not match the TOC's "TableOfFiguresLabel" property,
    // and is within the bounds of the bookmark. Its paragraph will not show up in the TOC as an entry.
    builder->Write(u"MySequence #");
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"OtherSequence");
    builder->Writeln(u", will not show up in the TOC because it's from a different sequence identifier.");
    
    // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bounds of the bookmark.
    // This field also references another bookmark. The contents of that bookmark will appear in the TOC entry for this SEQ field.
    // The SEQ field itself will not display the contents of that bookmark.
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    fieldSeq->set_BookmarkName(u"SEQBookmark");
    ASSERT_EQ(fieldSeq->GetFieldCode(), u" SEQ  MySequence SEQBookmark");
    
    // Create a bookmark with contents that will show up in the TOC entry due to the above SEQ field referencing it.
    builder->InsertBreak(BreakType::PageBreak);
    builder->StartBookmark(u"SEQBookmark");
    builder->Write(u"MySequence #");
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    fieldSeq->set_SequenceIdentifier(u"MySequence");
    builder->Writeln(u", text from inside SEQBookmark.");
    builder->EndBookmark(u"SEQBookmark");
    
    builder->EndBookmark(u"TOCBookmark");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.SEQ.Bookmark.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SEQ.Bookmark.docx"));
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 8);
    
    fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
    System::ArrayPtr<System::String> pageRefIds = fieldToc->get_Result().Split(System::MakeArray<char16_t>({u' '}))->LINQ_Where(static_cast<System::Func<System::String, bool>>(static_cast<std::function<bool(System::String s)>>([](System::String s) -> bool
    {
        return s.StartsWith(u"_Toc");
    })))->LINQ_ToArray();
    
    ASSERT_EQ(fieldToc->get_Type(), FieldType::FieldTOC);
    ASSERT_EQ(fieldToc->get_TableOfFiguresLabel(), u"MySequence");
    TestUtil::VerifyField(FieldType::FieldTOC, u" TOC  \\c MySequence \\b TOCBookmark", System::String::Format(u"MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF {0} \\h \u00142\u0015\r", pageRefIds[0]) + System::String::Format(u"3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF {0} \\h \u00142\u0015\r", pageRefIds[1]), fieldToc);
    
    auto fieldPageRef = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, System::String::Format(u" PAGEREF {0} \\h ", pageRefIds[0]), u"2", fieldPageRef);
    ASSERT_EQ(fieldPageRef->get_BookmarkName(), pageRefIds[0]);
    
    fieldPageRef = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, System::String::Format(u" PAGEREF {0} \\h ", pageRefIds[1]), u"2", fieldPageRef);
    ASSERT_EQ(fieldPageRef->get_BookmarkName(), pageRefIds[1]);
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"1", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"2", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(5));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  OtherSequence", u"1", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"OtherSequence");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(6));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence SEQBookmark", u"3", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
    ASSERT_EQ(fieldSeq->get_BookmarkName(), u"SEQBookmark");
    
    fieldSeq = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(doc->get_Range()->get_Fields()->idx_get(7));
    
    TestUtil::VerifyField(FieldType::FieldSequence, u" SEQ  MySequence", u"3", fieldSeq);
    ASSERT_EQ(fieldSeq->get_SequenceIdentifier(), u"MySequence");
}

namespace gtest_test
{

TEST_F(ExField, TocSeqBookmark)
{
    s_instance->TocSeqBookmark();
}

} // namespace gtest_test

void ExField::ChangeBibliographyStyles()
{
    auto oldCulture = System::Threading::Thread::get_CurrentThread()->get_CurrentCulture();
    //ExSkip
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::MakeObject<System::Globalization::CultureInfo>(u"en-nz", false));
    //ExSkip
    
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Bibliography.docx"));
    
    // If the document already has a style you can change it with the following code:
    // doc.Bibliography.BibliographyStyle = "Bibliography custom style.xsl";
    
    doc->get_FieldOptions()->set_BibliographyStylesProvider(System::MakeObject<Aspose::Words::ApiExamples::ExField::BibliographyStylesProvider>());
    doc->UpdateFields();
    
    doc->Save(get_ArtifactsDir() + u"Field.ChangeBibliographyStyles.docx");
    
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(oldCulture);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, ChangeBibliographyStyles)
{
    s_instance->ChangeBibliographyStyles();
}

} // namespace gtest_test

void ExField::FieldData()
{
    //ExStart
    //ExFor:FieldData
    //ExSummary:Shows how to insert a DATA field into a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldData>(builder->InsertField(FieldType::FieldData, true));
    ASSERT_EQ(field->GetFieldCode(), u" DATA ");
    //ExEnd
    
    TestUtil::VerifyField(FieldType::FieldData, u" DATA ", System::String::Empty, DocumentHelper::SaveOpen(doc)->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, FieldData)
{
    s_instance->FieldData();
}

} // namespace gtest_test

void ExField::FieldInclude()
{
    //ExStart
    //ExFor:FieldInclude
    //ExFor:FieldInclude.BookmarkName
    //ExFor:FieldInclude.LockFields
    //ExFor:FieldInclude.SourceFullName
    //ExFor:FieldInclude.TextConverter
    //ExSummary:Shows how to create an INCLUDE field, and set its properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // We can use an INCLUDE field to import a portion of another document in the local file system.
    // The bookmark from the other document that we reference with this field contains this imported portion.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldInclude>(builder->InsertField(FieldType::FieldInclude, true));
    field->set_SourceFullName(get_MyDir() + u"Bookmarks.docx");
    field->set_BookmarkName(u"MyBookmark1");
    field->set_LockFields(false);
    field->set_TextConverter(u"Microsoft Word");
    
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->GetFieldCode(), u" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"")->get_Success());
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.INCLUDE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INCLUDE.docx"));
    field = System::ExplicitCast<Aspose::Words::Fields::FieldInclude>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->get_Type(), FieldType::FieldInclude);
    ASSERT_EQ(field->get_Result(), u"First bookmark.");
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->GetFieldCode(), u" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"")->get_Success());
    
    ASSERT_EQ(field->get_SourceFullName(), get_MyDir() + u"Bookmarks.docx");
    ASSERT_EQ(field->get_BookmarkName(), u"MyBookmark1");
    ASSERT_FALSE(field->get_LockFields());
    ASSERT_EQ(field->get_TextConverter(), u"Microsoft Word");
}

namespace gtest_test
{

TEST_F(ExField, FieldInclude)
{
    s_instance->FieldInclude();
}

} // namespace gtest_test

void ExField::FieldIncludePicture()
{
    //ExStart
    //ExFor:FieldIncludePicture
    //ExFor:FieldIncludePicture.GraphicFilter
    //ExFor:FieldIncludePicture.IsLinked
    //ExFor:FieldIncludePicture.ResizeHorizontally
    //ExFor:FieldIncludePicture.ResizeVertically
    //ExFor:FieldIncludePicture.SourceFullName
    //ExFor:FieldImport
    //ExFor:FieldImport.GraphicFilter
    //ExFor:FieldImport.IsLinked
    //ExFor:FieldImport.SourceFullName
    //ExSummary:Shows how to insert images using IMPORT and INCLUDEPICTURE fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two similar field types that we can use to display images linked from the local file system.
    // 1 -  The INCLUDEPICTURE field:
    auto fieldIncludePicture = System::ExplicitCast<Aspose::Words::Fields::FieldIncludePicture>(builder->InsertField(FieldType::FieldIncludePicture, true));
    fieldIncludePicture->set_SourceFullName(get_ImageDir() + u"Transparent background logo.png");
    
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(fieldIncludePicture->GetFieldCode(), u" INCLUDEPICTURE  .*")->get_Success());
    
    // Apply the PNG32.FLT filter.
    fieldIncludePicture->set_GraphicFilter(u"PNG32");
    fieldIncludePicture->set_IsLinked(true);
    fieldIncludePicture->set_ResizeHorizontally(true);
    fieldIncludePicture->set_ResizeVertically(true);
    
    // 2 -  The IMPORT field:
    auto fieldImport = System::ExplicitCast<Aspose::Words::Fields::FieldImport>(builder->InsertField(FieldType::FieldImport, true));
    fieldImport->set_SourceFullName(get_ImageDir() + u"Transparent background logo.png");
    fieldImport->set_GraphicFilter(u"PNG32");
    fieldImport->set_IsLinked(true);
    
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(fieldImport->GetFieldCode(), u" IMPORT  .* \\\\c PNG32 \\\\d")->get_Success());
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.IMPORT.INCLUDEPICTURE.docx");
    //ExEnd
    
    ASSERT_EQ(fieldIncludePicture->get_SourceFullName(), get_ImageDir() + u"Transparent background logo.png");
    ASSERT_EQ(fieldIncludePicture->get_GraphicFilter(), u"PNG32");
    ASSERT_TRUE(fieldIncludePicture->get_IsLinked());
    ASSERT_TRUE(fieldIncludePicture->get_ResizeHorizontally());
    ASSERT_TRUE(fieldIncludePicture->get_ResizeVertically());
    
    ASSERT_EQ(fieldImport->get_SourceFullName(), get_ImageDir() + u"Transparent background logo.png");
    ASSERT_EQ(fieldImport->get_GraphicFilter(), u"PNG32");
    ASSERT_TRUE(fieldImport->get_IsLinked());
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.IMPORT.INCLUDEPICTURE.docx"));
    
    // The INCLUDEPICTURE fields have been converted into shapes with linked images during loading.
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 0);
    ASSERT_EQ(doc->GetChildNodes(NodeType::Shape, true)->get_Count(), 2);
    
    auto image = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(NodeType::Shape, 0, true));
    
    ASSERT_TRUE(image->get_IsImage());
    ASSERT_TRUE(System::TestTools::IsNull(image->get_ImageData()->get_ImageBytes()));
    ASSERT_EQ(image->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "), get_ImageDir() + u"Transparent background logo.png");
    
    image = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(NodeType::Shape, 1, true));
    
    ASSERT_TRUE(image->get_IsImage());
    ASSERT_TRUE(System::TestTools::IsNull(image->get_ImageData()->get_ImageBytes()));
    ASSERT_EQ(image->get_ImageData()->get_SourceFullName().Replace(u"%20", u" "), get_ImageDir() + u"Transparent background logo.png");
}

namespace gtest_test
{

TEST_F(ExField, FieldIncludePicture)
{
    s_instance->FieldIncludePicture();
}

} // namespace gtest_test

void ExField::FieldHyperlink()
{
    //ExStart
    //ExFor:FieldHyperlink
    //ExFor:FieldHyperlink.Address
    //ExFor:FieldHyperlink.IsImageMap
    //ExFor:FieldHyperlink.OpenInNewWindow
    //ExFor:FieldHyperlink.ScreenTip
    //ExFor:FieldHyperlink.SubAddress
    //ExFor:FieldHyperlink.Target
    //ExSummary:Shows how to use HYPERLINK fields to link to documents in the local file system.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(builder->InsertField(FieldType::FieldHyperlink, true));
    
    // When we click this HYPERLINK field in Microsoft Word,
    // it will open the linked document and then place the cursor at the specified bookmark.
    field->set_Address(get_MyDir() + u"Bookmarks.docx");
    field->set_SubAddress(u"MyBookmark3");
    field->set_ScreenTip(System::String(u"Open ") + field->get_Address() + u" on bookmark " + field->get_SubAddress() + u" in a new window");
    
    builder->Writeln();
    
    // When we click this HYPERLINK field in Microsoft Word,
    // it will open the linked document, and automatically scroll down to the specified iframe.
    field = System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(builder->InsertField(FieldType::FieldHyperlink, true));
    field->set_Address(get_MyDir() + u"Iframes.html");
    field->set_ScreenTip(System::String(u"Open ") + field->get_Address());
    field->set_Target(u"iframe_3");
    field->set_OpenInNewWindow(true);
    field->set_IsImageMap(false);
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.HYPERLINK.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.HYPERLINK.docx"));
    field = System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldHyperlink, System::String(u" HYPERLINK \"") + get_MyDir().Replace(u"\\", u"\\\\") + u"Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + get_MyDir() + u"Bookmarks.docx on bookmark MyBookmark3 in a new window\" ", get_MyDir() + u"Bookmarks.docx - MyBookmark3", field);
    ASSERT_EQ(field->get_Address(), get_MyDir() + u"Bookmarks.docx");
    ASSERT_EQ(field->get_SubAddress(), u"MyBookmark3");
    ASSERT_EQ(field->get_ScreenTip(), System::String(u"Open ") + field->get_Address().Replace(u"\\", System::String::Empty) + u" on bookmark " + field->get_SubAddress() + u" in a new window");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldHyperlink, System::String(u" HYPERLINK \"file:///") + get_MyDir().Replace(u"\\", u"\\\\").Replace(u" ", u"%20") + u"Iframes.html\" \\t \"iframe_3\" \\o \"Open " + get_MyDir().Replace(u"\\", u"\\\\") + u"Iframes.html\" ", get_MyDir() + u"Iframes.html", field);
    ASSERT_EQ(field->get_Address(), System::String(u"file:///") + get_MyDir().Replace(u" ", u"%20") + u"Iframes.html");
    ASSERT_EQ(field->get_ScreenTip(), System::String(u"Open ") + get_MyDir() + u"Iframes.html");
    ASSERT_EQ(field->get_Target(), u"iframe_3");
    ASSERT_FALSE(field->get_OpenInNewWindow());
    ASSERT_FALSE(field->get_IsImageMap());
}

namespace gtest_test
{

TEST_F(ExField, DISABLED_FieldHyperlink)
{
    s_instance->FieldHyperlink();
}

} // namespace gtest_test

void ExField::FieldIndexFilter()
{
    //ExStart
    //ExFor:FieldIndex
    //ExFor:FieldIndex.BookmarkName
    //ExFor:FieldIndex.EntryType
    //ExFor:FieldXE
    //ExFor:FieldXE.EntryType
    //ExFor:FieldXE.Text
    //ExSummary:Shows how to create an INDEX field, and then use XE fields to populate it with entries.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side
    // and the page containing the XE field on the right.
    // If the XE fields have the same value in their "Text" property,
    // the INDEX field will group them into one entry.
    auto index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
    
    // Configure the INDEX field only to display XE fields that are within the bounds
    // of a bookmark named "MainBookmark", and whose "EntryType" properties have a value of "A".
    // For both INDEX and XE fields, the "EntryType" property only uses the first character of its string value.
    index->set_BookmarkName(u"MainBookmark");
    index->set_EntryType(u"A");
    
    ASSERT_EQ(index->GetFieldCode(), u" INDEX  \\b MainBookmark \\f A");
    
    // On a new page, start the bookmark with a name that matches the value
    // of the INDEX field's "BookmarkName" property.
    builder->InsertBreak(BreakType::PageBreak);
    builder->StartBookmark(u"MainBookmark");
    
    // The INDEX field will pick up this entry because it is inside the bookmark,
    // and its entry type also matches the INDEX field's entry type.
    auto indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Index entry 1");
    indexEntry->set_EntryType(u"A");
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  \"Index entry 1\" \\f A");
    
    // Insert an XE field that will not appear in the INDEX because the entry types do not match.
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Index entry 2");
    indexEntry->set_EntryType(u"B");
    
    // End the bookmark and insert an XE field afterwards.
    // It is of the same type as the INDEX field, but will not appear
    // since it is outside the bookmark's boundaries.
    builder->EndBookmark(u"MainBookmark");
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Index entry 3");
    indexEntry->set_EntryType(u"A");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.INDEX.XE.Filtering.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INDEX.XE.Filtering.docx"));
    index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\b MainBookmark \\f A", u"Index entry 1, 2\r", index);
    ASSERT_EQ(index->get_BookmarkName(), u"MainBookmark");
    ASSERT_EQ(index->get_EntryType(), u"A");
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Index entry 1\" \\f A", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Index entry 1");
    ASSERT_EQ(indexEntry->get_EntryType(), u"A");
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Index entry 2\" \\f B", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Index entry 2");
    ASSERT_EQ(indexEntry->get_EntryType(), u"B");
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Index entry 3\" \\f A", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Index entry 3");
    ASSERT_EQ(indexEntry->get_EntryType(), u"A");
}

namespace gtest_test
{

TEST_F(ExField, FieldIndexFilter)
{
    s_instance->FieldIndexFilter();
}

} // namespace gtest_test

void ExField::FieldIndexFormatting()
{
    //ExStart
    //ExFor:FieldIndex
    //ExFor:FieldIndex.Heading
    //ExFor:FieldIndex.NumberOfColumns
    //ExFor:FieldIndex.LanguageId
    //ExFor:FieldIndex.LetterRange
    //ExFor:FieldXE
    //ExFor:FieldXE.IsBold
    //ExFor:FieldXE.IsItalic
    //ExFor:FieldXE.Text
    //ExSummary:Shows how to populate an INDEX field with entries using XE fields, and also modify its appearance.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // If the XE fields have the same value in their "Text" property,
    // the INDEX field will group them into one entry.
    auto index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
    index->set_LanguageId(u"1033");
    
    // Setting this property's value to "A" will group all the entries by their first letter,
    // and place that letter in uppercase above each group.
    index->set_Heading(u"A");
    
    // Set the table created by the INDEX field to span over 2 columns.
    index->set_NumberOfColumns(u"2");
    
    // Set any entries with starting letters outside the "a-c" character range to be omitted.
    index->set_LetterRange(u"a-c");
    
    ASSERT_EQ(index->GetFieldCode(), u" INDEX  \\z 1033 \\h A \\c 2 \\p a-c");
    
    // These next two XE fields will show up under the "A" heading,
    // with their respective text stylings also applied to their page numbers.
    builder->InsertBreak(BreakType::PageBreak);
    auto indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Apple");
    indexEntry->set_IsItalic(true);
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  Apple \\i");
    
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Apricot");
    indexEntry->set_IsBold(true);
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  Apricot \\b");
    
    // Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents.
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Banana");
    
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Cherry");
    
    // INDEX fields sort all entries alphabetically, so this entry will show up under "A" with the other two.
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Avocado");
    
    // This entry will not appear because it starts with the letter "D",
    // which is outside the "a-c" character range that the INDEX field's LetterRange property defines.
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Durian");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.INDEX.XE.Formatting.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INDEX.XE.Formatting.docx"));
    index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(index->get_LanguageId(), u"1033");
    ASSERT_EQ(index->get_Heading(), u"A");
    ASSERT_EQ(index->get_NumberOfColumns(), u"2");
    ASSERT_EQ(index->get_LetterRange(), u"a-c");
    ASSERT_EQ(index->GetFieldCode(), u" INDEX  \\z 1033 \\h A \\c 2 \\p a-c");
    ASSERT_EQ(index->get_Result(), (System::String(u"\fA\r") + u"Apple, 2\r" + u"Apricot, 3\r" + u"Avocado, 6\r" + u"B\r" + u"Banana, 4\r" + u"C\r" + u"Cherry, 5\r\f"));
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Apple \\i", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Apple");
    ASSERT_FALSE(indexEntry->get_IsBold());
    ASSERT_TRUE(indexEntry->get_IsItalic());
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Apricot \\b", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Apricot");
    ASSERT_TRUE(indexEntry->get_IsBold());
    ASSERT_FALSE(indexEntry->get_IsItalic());
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Banana", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Banana");
    ASSERT_FALSE(indexEntry->get_IsBold());
    ASSERT_FALSE(indexEntry->get_IsItalic());
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Cherry", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Cherry");
    ASSERT_FALSE(indexEntry->get_IsBold());
    ASSERT_FALSE(indexEntry->get_IsItalic());
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(5));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Avocado", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Avocado");
    ASSERT_FALSE(indexEntry->get_IsBold());
    ASSERT_FALSE(indexEntry->get_IsItalic());
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(6));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Durian", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Durian");
    ASSERT_FALSE(indexEntry->get_IsBold());
    ASSERT_FALSE(indexEntry->get_IsItalic());
}

namespace gtest_test
{

TEST_F(ExField, FieldIndexFormatting)
{
    s_instance->FieldIndexFormatting();
}

} // namespace gtest_test

void ExField::FieldIndexSequence()
{
    //ExStart
    //ExFor:FieldIndex.HasSequenceName
    //ExFor:FieldIndex.SequenceName
    //ExFor:FieldIndex.SequenceSeparator
    //ExSummary:Shows how to split a document into portions by combining INDEX and SEQ fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // If the XE fields have the same value in their "Text" property,
    // the INDEX field will group them into one entry.
    auto index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
    
    // In the SequenceName property, name a SEQ field sequence. Each entry of this INDEX field will now also display
    // the number that the sequence count is on at the XE field location that created this entry.
    index->set_SequenceName(u"MySequence");
    
    // Set text that will around the sequence and page numbers to explain their meaning to the user.
    // An entry created with this configuration will display something like "MySequence at 1 on page 1" at its page number.
    // PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters.
    index->set_PageNumberSeparator(u"\tMySequence at ");
    index->set_SequenceSeparator(u" on page ");
    ASSERT_TRUE(index->get_HasSequenceName());
    
    ASSERT_EQ(index->GetFieldCode(), u" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"");
    
    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Insert a SEQ field which moves the "MySequence" sequence to 1.
    // This field no different from normal document text. It will not appear on an INDEX field's table of contents.
    builder->InsertBreak(BreakType::PageBreak);
    auto sequenceField = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    sequenceField->set_SequenceIdentifier(u"MySequence");
    
    ASSERT_EQ(sequenceField->GetFieldCode(), u" SEQ  MySequence");
    
    // Insert an XE field which will create an entry in the INDEX field.
    // Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
    // this field's INDEX entry will display "Cat" on the left side, and "MySequence at 1 on page 2" on the right.
    auto indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Cat");
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  Cat");
    
    // Insert a page break and use SEQ fields to advance "MySequence" to 3.
    builder->InsertBreak(BreakType::PageBreak);
    sequenceField = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    sequenceField->set_SequenceIdentifier(u"MySequence");
    sequenceField = System::ExplicitCast<Aspose::Words::Fields::FieldSeq>(builder->InsertField(FieldType::FieldSequence, true));
    sequenceField->set_SequenceIdentifier(u"MySequence");
    
    // Insert an XE field with the same Text property as the one above.
    // The INDEX entry will group XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    // Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above.
    // The page number portion of that INDEX entry will now display "MySequence at 1 on page 2, 3 on page 3".
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Cat");
    
    // Insert an XE field with a new and unique Text property value.
    // This will add a new entry, with MySequence at 3 on page 4.
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Dog");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.INDEX.XE.Sequence.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INDEX.XE.Sequence.docx"));
    index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(index->get_SequenceName(), u"MySequence");
    ASSERT_EQ(index->get_PageNumberSeparator(), u"\tMySequence at ");
    ASSERT_EQ(index->get_SequenceSeparator(), u" on page ");
    ASSERT_TRUE(index->get_HasSequenceName());
    ASSERT_EQ(index->GetFieldCode(), u" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"");
    ASSERT_EQ(index->get_Result(), System::String(u"Cat\tMySequence at 1 on page 2, 3 on page 3\r") + u"Dog\tMySequence at 3 on page 4\r");
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->LINQ_Where(static_cast<System::Func<System::SharedPtr<Aspose::Words::Fields::Field>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Fields::Field> f)>>([](System::SharedPtr<Aspose::Words::Fields::Field> f) -> bool
    {
        return f->get_Type() == FieldType::FieldSequence;
    })))->LINQ_Count(), 3);
}

namespace gtest_test
{

TEST_F(ExField, FieldIndexSequence)
{
    s_instance->FieldIndexSequence();
}

} // namespace gtest_test

void ExField::FieldIndexPageNumberSeparator()
{
    //ExStart
    //ExFor:FieldIndex.HasPageNumberSeparator
    //ExFor:FieldIndex.PageNumberSeparator
    //ExFor:FieldIndex.PageNumberListSeparator
    //ExSummary:Shows how to edit the page number separator in an INDEX field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will group XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    auto index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
    
    // If our INDEX field has an entry for a group of XE fields,
    // this entry will display the number of each page that contains an XE field that belongs to this group.
    // We can set custom separators to customize the appearance of these page numbers.
    index->set_PageNumberSeparator(u", on page(s) ");
    index->set_PageNumberListSeparator(u" & ");
    
    ASSERT_EQ(index->GetFieldCode(), (u" INDEX  \\e \", on page(s) \" \\l \" & \""));
    ASSERT_TRUE(index->get_HasPageNumberSeparator());
    
    // After we insert these XE fields, the INDEX field will display "First entry, on page(s) 2 & 3 & 4".
    builder->InsertBreak(BreakType::PageBreak);
    auto indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"First entry");
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  \"First entry\"");
    
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"First entry");
    
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"First entry");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.INDEX.XE.PageNumberList.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INDEX.XE.PageNumberList.docx"));
    index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\e \", on page(s) \" \\l \" & \"", u"First entry, on page(s) 2 & 3 & 4\r", index);
    ASSERT_EQ(index->get_PageNumberSeparator(), (u", on page(s) "));
    ASSERT_EQ(index->get_PageNumberListSeparator(), u" & ");
    ASSERT_TRUE(index->get_HasPageNumberSeparator());
}

namespace gtest_test
{

TEST_F(ExField, FieldIndexPageNumberSeparator)
{
    s_instance->FieldIndexPageNumberSeparator();
}

} // namespace gtest_test

void ExField::FieldIndexPageRangeBookmark()
{
    //ExStart
    //ExFor:FieldIndex.PageRangeSeparator
    //ExFor:FieldXE.PageRangeBookmarkName
    //ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will collect all XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    auto index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
    
    // For INDEX entries that display page ranges, we can specify a separator string
    // which will appear between the number of the first page, and the number of the last.
    index->set_PageNumberSeparator(u", on page(s) ");
    index->set_PageRangeSeparator(u" to ");
    
    ASSERT_EQ(index->GetFieldCode(), (u" INDEX  \\e \", on page(s) \" \\g \" to \""));
    
    builder->InsertBreak(BreakType::PageBreak);
    auto indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"My entry");
    
    // If an XE field names a bookmark using the PageRangeBookmarkName property,
    // its INDEX entry will show the range of pages that the bookmark spans
    // instead of the number of the page that contains the XE field.
    indexEntry->set_PageRangeBookmarkName(u"MyBookmark");
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  \"My entry\" \\r MyBookmark");
    ASSERT_EQ(indexEntry->get_PageRangeBookmarkName(), u"MyBookmark");
    
    // Insert a bookmark that starts on page 3 and ends on page 5.
    // The INDEX entry for the XE field that references this bookmark will display this page range.
    // In our table, the INDEX entry will display "My entry, on page(s) 3 to 5".
    builder->InsertBreak(BreakType::PageBreak);
    builder->StartBookmark(u"MyBookmark");
    builder->Write(u"Start of MyBookmark");
    builder->InsertBreak(BreakType::PageBreak);
    builder->InsertBreak(BreakType::PageBreak);
    builder->Write(u"End of MyBookmark");
    builder->EndBookmark(u"MyBookmark");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.INDEX.XE.PageRangeBookmark.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INDEX.XE.PageRangeBookmark.docx"));
    index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\e \", on page(s) \" \\g \" to \"", u"My entry, on page(s) 3 to 5\r", index);
    ASSERT_EQ(index->get_PageNumberSeparator(), (u", on page(s) "));
    ASSERT_EQ(index->get_PageRangeSeparator(), u" to ");
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"My entry\" \\r MyBookmark", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"My entry");
    ASSERT_EQ(indexEntry->get_PageRangeBookmarkName(), u"MyBookmark");
}

namespace gtest_test
{

TEST_F(ExField, FieldIndexPageRangeBookmark)
{
    s_instance->FieldIndexPageRangeBookmark();
}

} // namespace gtest_test

void ExField::FieldIndexCrossReferenceSeparator()
{
    //ExStart
    //ExFor:FieldIndex.CrossReferenceSeparator
    //ExFor:FieldXE.PageNumberReplacement
    //ExSummary:Shows how to define cross references in an INDEX field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will collect all XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    auto index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
    
    // We can configure an XE field to get its INDEX entry to display a string instead of a page number.
    // First, for entries that substitute a page number with a string,
    // specify a custom separator between the XE field's Text property value and the string.
    index->set_CrossReferenceSeparator(u", see: ");
    
    ASSERT_EQ(index->GetFieldCode(), (u" INDEX  \\k \", see: \""));
    
    // Insert an XE field, which creates a regular INDEX entry which displays this field's page number,
    // and does not invoke the CrossReferenceSeparator value.
    // The entry for this XE field will display "Apple, 2".
    builder->InsertBreak(BreakType::PageBreak);
    auto indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Apple");
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  Apple");
    
    // Insert another XE field on page 3 and set a value for the PageNumberReplacement property.
    // This value will show up instead of the number of the page that this field is on,
    // and the INDEX field's CrossReferenceSeparator value will appear in front of it.
    // The entry for this XE field will display "Banana, see: Tropical fruit".
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Banana");
    indexEntry->set_PageNumberReplacement(u"Tropical fruit");
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  Banana \\t \"Tropical fruit\"");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.INDEX.XE.CrossReferenceSeparator.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INDEX.XE.CrossReferenceSeparator.docx"));
    index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\k \", see: \"", System::String(u"Apple, 2\r") + u"Banana, see: Tropical fruit\r", index);
    ASSERT_EQ(index->get_CrossReferenceSeparator(), (u", see: "));
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Apple", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Apple");
    ASSERT_TRUE(System::TestTools::IsNull(indexEntry->get_PageNumberReplacement()));
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  Banana \\t \"Tropical fruit\"", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Banana");
    ASSERT_EQ(indexEntry->get_PageNumberReplacement(), u"Tropical fruit");
}

namespace gtest_test
{

TEST_F(ExField, FieldIndexCrossReferenceSeparator)
{
    s_instance->FieldIndexCrossReferenceSeparator();
}

} // namespace gtest_test

void ExField::FieldIndexSubheading(bool runSubentriesOnTheSameLine)
{
    //ExStart
    //ExFor:FieldIndex.RunSubentriesOnSameLine
    //ExSummary:Shows how to work with subentries in an INDEX field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will collect all XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    auto index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(builder->InsertField(FieldType::FieldIndex, true));
    index->set_PageNumberSeparator(u", see page ");
    index->set_Heading(u"A");
    
    // XE fields that have a Text property whose value becomes the heading of the INDEX entry.
    // If this value contains two string segments split by a colon (the INDEX entry will treat :) delimiter,
    // the first segment is heading, and the second segment will become the subheading.
    // The INDEX field first groups entries alphabetically, then, if there are multiple XE fields with the same
    // headings, the INDEX field will further subgroup them by the values of these headings.
    // There can be multiple subgrouping layers, depending on how many times
    // the Text properties of XE fields get segmented like this.
    // By default, an INDEX field entry group will create a new line for every subheading within this group.
    // We can set the RunSubentriesOnSameLine flag to true to keep the heading,
    // and every subheading for the group on one line instead, which will make the INDEX field more compact.
    index->set_RunSubentriesOnSameLine(runSubentriesOnTheSameLine);
    
    if (runSubentriesOnTheSameLine)
    {
        ASSERT_EQ(index->GetFieldCode(), (u" INDEX  \\e \", see page \" \\h A \\r"));
    }
    else
    {
        ASSERT_EQ(index->GetFieldCode(), (u" INDEX  \\e \", see page \" \\h A"));
    }
    
    // Insert two XE fields, each on a new page, and with the same heading named "Heading 1",
    // which the INDEX field will use to group them.
    // If RunSubentriesOnSameLine is false, then the INDEX table will create three lines:
    // one line for the grouping heading "Heading 1", and one more line for each subheading.
    // If RunSubentriesOnSameLine is true, then the INDEX table will create a one-line
    // entry that encompasses the heading and every subheading.
    builder->InsertBreak(BreakType::PageBreak);
    auto indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Heading 1:Subheading 1");
    
    ASSERT_EQ(indexEntry->GetFieldCode(), u" XE  \"Heading 1:Subheading 1\"");
    
    builder->InsertBreak(BreakType::PageBreak);
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(builder->InsertField(FieldType::FieldIndexEntry, true));
    indexEntry->set_Text(u"Heading 1:Subheading 2");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + System::String::Format(u"Field.INDEX.XE.Subheading.docx", System::MakeObject<System::Array<System::SharedPtr<System::Object>>>(0)));
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + System::String::Format(u"Field.INDEX.XE.Subheading.docx", System::MakeObject<System::Array<System::SharedPtr<System::Object>>>(0))));
    index = System::ExplicitCast<Aspose::Words::Fields::FieldIndex>(doc->get_Range()->get_Fields()->idx_get(0));
    
    if (runSubentriesOnTheSameLine)
    {
        TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\e \", see page \" \\h A \\r", System::String(u"H\r") + u"Heading 1: Subheading 1, see page 2; Subheading 2, see page 3\r", index);
        ASSERT_TRUE(index->get_RunSubentriesOnSameLine());
    }
    else
    {
        TestUtil::VerifyField(FieldType::FieldIndex, u" INDEX  \\e \", see page \" \\h A", System::String(u"H\r") + u"Heading 1\r" + u"Subheading 1, see page 2\r" + u"Subheading 2, see page 3\r", index);
        ASSERT_FALSE(index->get_RunSubentriesOnSameLine());
    }
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Heading 1:Subheading 1\"", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Heading 1:Subheading 1");
    
    indexEntry = System::ExplicitCast<Aspose::Words::Fields::FieldXE>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldIndexEntry, u" XE  \"Heading 1:Subheading 2\"", System::String::Empty, indexEntry);
    ASSERT_EQ(indexEntry->get_Text(), u"Heading 1:Subheading 2");
}

namespace gtest_test
{

using ExField_FieldIndexSubheading_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExField::FieldIndexSubheading)>::type;

struct ExField_FieldIndexSubheading : public ExField, public Aspose::Words::ApiExamples::ExField, public ::testing::WithParamInterface<ExField_FieldIndexSubheading_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ExField_FieldIndexSubheading, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FieldIndexSubheading(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_FieldIndexSubheading, ::testing::ValuesIn(ExField_FieldIndexSubheading::TestCases()));

} // namespace gtest_test

void ExField::FieldBarcode()
{
    //ExStart
    //ExFor:FieldBarcode
    //ExFor:FieldBarcode.FacingIdentificationMark
    //ExFor:FieldBarcode.IsBookmark
    //ExFor:FieldBarcode.IsUSPostalAddress
    //ExFor:FieldBarcode.PostalAddress
    //ExSummary:Shows how to use the BARCODE field to display U.S. ZIP codes in the form of a barcode.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Writeln();
    
    // Below are two ways of using BARCODE fields to display custom values as barcodes.
    // 1 -  Store the value that the barcode will display in the PostalAddress property:
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldBarcode>(builder->InsertField(FieldType::FieldBarcode, true));
    
    // This value needs to be a valid ZIP code.
    field->set_PostalAddress(u"96801");
    field->set_IsUSPostalAddress(true);
    field->set_FacingIdentificationMark(u"C");
    
    ASSERT_EQ(field->GetFieldCode(), u" BARCODE  96801 \\u \\f C");
    
    builder->InsertBreak(BreakType::LineBreak);
    
    // 2 -  Reference a bookmark that stores the value that this barcode will display:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldBarcode>(builder->InsertField(FieldType::FieldBarcode, true));
    field->set_PostalAddress(u"BarcodeBookmark");
    field->set_IsBookmark(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" BARCODE  BarcodeBookmark \\b");
    
    // The bookmark that the BARCODE field references in its PostalAddress property
    // need to contain nothing besides the valid ZIP code.
    builder->InsertBreak(BreakType::PageBreak);
    builder->StartBookmark(u"BarcodeBookmark");
    builder->Writeln(u"968877");
    builder->EndBookmark(u"BarcodeBookmark");
    
    doc->Save(get_ArtifactsDir() + u"Field.BARCODE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.BARCODE.docx"));
    
    ASSERT_EQ(doc->GetChildNodes(NodeType::Shape, true)->get_Count(), 0);
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldBarcode>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE  96801 \\u \\f C", System::String::Empty, field);
    ASSERT_EQ(field->get_FacingIdentificationMark(), u"C");
    ASSERT_EQ(field->get_PostalAddress(), u"96801");
    ASSERT_TRUE(field->get_IsUSPostalAddress());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldBarcode>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldBarcode, u" BARCODE  BarcodeBookmark \\b", System::String::Empty, field);
    ASSERT_EQ(field->get_PostalAddress(), u"BarcodeBookmark");
    ASSERT_TRUE(field->get_IsBookmark());
}

namespace gtest_test
{

TEST_F(ExField, FieldBarcode)
{
    s_instance->FieldBarcode();
}

} // namespace gtest_test

void ExField::FieldDisplayBarcode()
{
    //ExStart
    //ExFor:FieldDisplayBarcode
    //ExFor:FieldDisplayBarcode.AddStartStopChar
    //ExFor:FieldDisplayBarcode.BackgroundColor
    //ExFor:FieldDisplayBarcode.BarcodeType
    //ExFor:FieldDisplayBarcode.BarcodeValue
    //ExFor:FieldDisplayBarcode.CaseCodeStyle
    //ExFor:FieldDisplayBarcode.DisplayText
    //ExFor:FieldDisplayBarcode.ErrorCorrectionLevel
    //ExFor:FieldDisplayBarcode.FixCheckDigit
    //ExFor:FieldDisplayBarcode.ForegroundColor
    //ExFor:FieldDisplayBarcode.PosCodeStyle
    //ExFor:FieldDisplayBarcode.ScalingFactor
    //ExFor:FieldDisplayBarcode.SymbolHeight
    //ExFor:FieldDisplayBarcode.SymbolRotation
    //ExSummary:Shows how to insert a DISPLAYBARCODE field, and set its properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
    
    // Below are four types of barcodes, decorated in various ways, that the DISPLAYBARCODE field can display.
    // 1 -  QR code with custom colors:
    field->set_BarcodeType(u"QR");
    field->set_BarcodeValue(u"ABC123");
    field->set_BackgroundColor(u"0xF8BD69");
    field->set_ForegroundColor(u"0xB5413B");
    field->set_ErrorCorrectionLevel(u"3");
    field->set_ScalingFactor(u"250");
    field->set_SymbolHeight(u"1000");
    field->set_SymbolRotation(u"0");
    
    ASSERT_EQ(field->GetFieldCode(), u" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0");
    builder->Writeln();
    
    // 2 -  EAN13 barcode, with the digits displayed below the bars:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
    field->set_BarcodeType(u"EAN13");
    field->set_BarcodeValue(u"501234567890");
    field->set_DisplayText(true);
    field->set_PosCodeStyle(u"CASE");
    field->set_FixCheckDigit(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x");
    builder->Writeln();
    
    // 3 -  CODE39 barcode:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
    field->set_BarcodeType(u"CODE39");
    field->set_BarcodeValue(u"12345ABCDE");
    field->set_AddStartStopChar(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" DISPLAYBARCODE  12345ABCDE CODE39 \\d");
    builder->Writeln();
    
    // 4 -  ITF4 barcode, with a specified case code:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(builder->InsertField(FieldType::FieldDisplayBarcode, true));
    field->set_BarcodeType(u"ITF14");
    field->set_BarcodeValue(u"09312345678907");
    field->set_CaseCodeStyle(u"STD");
    
    ASSERT_EQ(field->GetFieldCode(), u" DISPLAYBARCODE  09312345678907 ITF14 \\c STD");
    
    doc->Save(get_ArtifactsDir() + u"Field.DISPLAYBARCODE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.DISPLAYBARCODE.docx"));
    
    ASSERT_EQ(doc->GetChildNodes(NodeType::Shape, true)->get_Count(), 0);
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", System::String::Empty, field);
    ASSERT_EQ(field->get_BarcodeType(), u"QR");
    ASSERT_EQ(field->get_BarcodeValue(), u"ABC123");
    ASSERT_EQ(field->get_BackgroundColor(), u"0xF8BD69");
    ASSERT_EQ(field->get_ForegroundColor(), u"0xB5413B");
    ASSERT_EQ(field->get_ErrorCorrectionLevel(), u"3");
    ASSERT_EQ(field->get_ScalingFactor(), u"250");
    ASSERT_EQ(field->get_SymbolHeight(), u"1000");
    ASSERT_EQ(field->get_SymbolRotation(), u"0");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", System::String::Empty, field);
    ASSERT_EQ(field->get_BarcodeType(), u"EAN13");
    ASSERT_EQ(field->get_BarcodeValue(), u"501234567890");
    ASSERT_TRUE(field->get_DisplayText());
    ASSERT_EQ(field->get_PosCodeStyle(), u"CASE");
    ASSERT_TRUE(field->get_FixCheckDigit());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  12345ABCDE CODE39 \\d", System::String::Empty, field);
    ASSERT_EQ(field->get_BarcodeType(), u"CODE39");
    ASSERT_EQ(field->get_BarcodeValue(), u"12345ABCDE");
    ASSERT_TRUE(field->get_AddStartStopChar());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDisplayBarcode>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldDisplayBarcode, u" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", System::String::Empty, field);
    ASSERT_EQ(field->get_BarcodeType(), u"ITF14");
    ASSERT_EQ(field->get_BarcodeValue(), u"09312345678907");
    ASSERT_EQ(field->get_CaseCodeStyle(), u"STD");
}

namespace gtest_test
{

TEST_F(ExField, FieldDisplayBarcode)
{
    s_instance->FieldDisplayBarcode();
}

} // namespace gtest_test

void ExField::FieldLinkedObjectsAsText(Aspose::Words::ApiExamples::InsertLinkedObjectAs insertLinkedObjectAs)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are three types of fields we can use to display contents from a linked document in the form of text.
    // 1 -  A LINK field:
    builder->Writeln(u"FieldLink:\n");
    InsertFieldLink(builder, insertLinkedObjectAs, u"Word.Document.8", get_MyDir() + u"Document.docx", nullptr, true);
    
    // 2 -  A DDE field:
    builder->Writeln(u"FieldDde:\n");
    InsertFieldDde(builder, insertLinkedObjectAs, u"Excel.Sheet", get_MyDir() + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true, true);
    
    // 3 -  A DDEAUTO field:
    builder->Writeln(u"FieldDdeAuto:\n");
    InsertFieldDdeAuto(builder, insertLinkedObjectAs, u"Excel.Sheet", get_MyDir() + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true);
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.LINK.DDE.DDEAUTO.docx");
}

namespace gtest_test
{

using ExField_FieldLinkedObjectsAsText_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExField::FieldLinkedObjectsAsText)>::type;

struct ExField_FieldLinkedObjectsAsText : public ExField, public Aspose::Words::ApiExamples::ExField, public ::testing::WithParamInterface<ExField_FieldLinkedObjectsAsText_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(InsertLinkedObjectAs::Text),
            std::make_tuple(InsertLinkedObjectAs::Unicode),
            std::make_tuple(InsertLinkedObjectAs::Html),
            std::make_tuple(InsertLinkedObjectAs::Rtf),
        };
    }
};

TEST_P(ExField_FieldLinkedObjectsAsText, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FieldLinkedObjectsAsText(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_FieldLinkedObjectsAsText, ::testing::ValuesIn(ExField_FieldLinkedObjectsAsText::TestCases()));

} // namespace gtest_test

void ExField::FieldLinkedObjectsAsImage(Aspose::Words::ApiExamples::InsertLinkedObjectAs insertLinkedObjectAs)
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are three types of fields we can use to display contents from a linked document in the form of an image.
    // 1 -  A LINK field:
    builder->Writeln(u"FieldLink:\n");
    InsertFieldLink(builder, insertLinkedObjectAs, u"Excel.Sheet", get_MyDir() + u"MySpreadsheet.xlsx", u"Sheet1!R2C2", true);
    
    // 2 -  A DDE field:
    builder->Writeln(u"FieldDde:\n");
    InsertFieldDde(builder, insertLinkedObjectAs, u"Excel.Sheet", get_MyDir() + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true, true);
    
    // 3 -  A DDEAUTO field:
    builder->Writeln(u"FieldDdeAuto:\n");
    InsertFieldDdeAuto(builder, insertLinkedObjectAs, u"Excel.Sheet", get_MyDir() + u"Spreadsheet.xlsx", u"Sheet1!R1C1", true);
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.LINK.DDE.DDEAUTO.AsImage.docx");
}

namespace gtest_test
{

using ExField_FieldLinkedObjectsAsImage_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExField::FieldLinkedObjectsAsImage)>::type;

struct ExField_FieldLinkedObjectsAsImage : public ExField, public Aspose::Words::ApiExamples::ExField, public ::testing::WithParamInterface<ExField_FieldLinkedObjectsAsImage_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(InsertLinkedObjectAs::Picture),
            std::make_tuple(InsertLinkedObjectAs::Bitmap),
        };
    }
};

TEST_P(ExField_FieldLinkedObjectsAsImage, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->FieldLinkedObjectsAsImage(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_FieldLinkedObjectsAsImage, ::testing::ValuesIn(ExField_FieldLinkedObjectsAsImage::TestCases()));

} // namespace gtest_test

void ExField::FieldUserAddress()
{
    //ExStart
    //ExFor:FieldUserAddress
    //ExFor:FieldUserAddress.UserAddress
    //ExSummary:Shows how to use the USERADDRESS field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a UserInformation object and set it as the source of user information for any fields that we create.
    auto userInformation = System::MakeObject<Aspose::Words::Fields::UserInformation>();
    userInformation->set_Address(u"123 Main Street");
    doc->get_FieldOptions()->set_CurrentUser(userInformation);
    
    // Create a USERADDRESS field to display the current user's address,
    // taken from the UserInformation object we created above.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto fieldUserAddress = System::ExplicitCast<Aspose::Words::Fields::FieldUserAddress>(builder->InsertField(FieldType::FieldUserAddress, true));
    ASSERT_EQ(fieldUserAddress->get_Result(), userInformation->get_Address());
    //ExSkip
    
    ASSERT_EQ(fieldUserAddress->GetFieldCode(), u" USERADDRESS ");
    ASSERT_EQ(fieldUserAddress->get_Result(), u"123 Main Street");
    
    // We can set this property to get our field to override the value currently stored in the UserInformation object.
    fieldUserAddress->set_UserAddress(u"456 North Road");
    fieldUserAddress->Update();
    
    ASSERT_EQ(fieldUserAddress->GetFieldCode(), u" USERADDRESS  \"456 North Road\"");
    ASSERT_EQ(fieldUserAddress->get_Result(), u"456 North Road");
    
    // This does not affect the value in the UserInformation object.
    ASSERT_EQ(doc->get_FieldOptions()->get_CurrentUser()->get_Address(), u"123 Main Street");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.USERADDRESS.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.USERADDRESS.docx"));
    
    fieldUserAddress = System::ExplicitCast<Aspose::Words::Fields::FieldUserAddress>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldUserAddress, u" USERADDRESS  \"456 North Road\"", u"456 North Road", fieldUserAddress);
    ASSERT_EQ(fieldUserAddress->get_UserAddress(), u"456 North Road");
}

namespace gtest_test
{

TEST_F(ExField, FieldUserAddress)
{
    s_instance->FieldUserAddress();
}

} // namespace gtest_test

void ExField::FieldUserInitials()
{
    //ExStart
    //ExFor:FieldUserInitials
    //ExFor:FieldUserInitials.UserInitials
    //ExSummary:Shows how to use the USERINITIALS field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a UserInformation object and set it as the source of user information for any fields that we create.
    auto userInformation = System::MakeObject<Aspose::Words::Fields::UserInformation>();
    userInformation->set_Initials(u"J. D.");
    doc->get_FieldOptions()->set_CurrentUser(userInformation);
    
    // Create a USERINITIALS field to display the current user's initials,
    // taken from the UserInformation object we created above.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto fieldUserInitials = System::ExplicitCast<Aspose::Words::Fields::FieldUserInitials>(builder->InsertField(FieldType::FieldUserInitials, true));
    ASSERT_EQ(fieldUserInitials->get_Result(), userInformation->get_Initials());
    
    ASSERT_EQ(fieldUserInitials->GetFieldCode(), u" USERINITIALS ");
    ASSERT_EQ(fieldUserInitials->get_Result(), u"J. D.");
    
    // We can set this property to get our field to override the value currently stored in the UserInformation object.
    fieldUserInitials->set_UserInitials(u"J. C.");
    fieldUserInitials->Update();
    
    ASSERT_EQ(fieldUserInitials->GetFieldCode(), u" USERINITIALS  \"J. C.\"");
    ASSERT_EQ(fieldUserInitials->get_Result(), u"J. C.");
    
    // This does not affect the value in the UserInformation object.
    ASSERT_EQ(doc->get_FieldOptions()->get_CurrentUser()->get_Initials(), u"J. D.");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.USERINITIALS.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.USERINITIALS.docx"));
    
    fieldUserInitials = System::ExplicitCast<Aspose::Words::Fields::FieldUserInitials>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldUserInitials, u" USERINITIALS  \"J. C.\"", u"J. C.", fieldUserInitials);
    ASSERT_EQ(fieldUserInitials->get_UserInitials(), u"J. C.");
}

namespace gtest_test
{

TEST_F(ExField, FieldUserInitials)
{
    s_instance->FieldUserInitials();
}

} // namespace gtest_test

void ExField::FieldUserName()
{
    //ExStart
    //ExFor:FieldUserName
    //ExFor:FieldUserName.UserName
    //ExSummary:Shows how to use the USERNAME field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Create a UserInformation object and set it as the source of user information for any fields that we create.
    auto userInformation = System::MakeObject<Aspose::Words::Fields::UserInformation>();
    userInformation->set_Name(u"John Doe");
    doc->get_FieldOptions()->set_CurrentUser(userInformation);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a USERNAME field to display the current user's name,
    // taken from the UserInformation object we created above.
    auto fieldUserName = System::ExplicitCast<Aspose::Words::Fields::FieldUserName>(builder->InsertField(FieldType::FieldUserName, true));
    ASSERT_EQ(fieldUserName->get_Result(), userInformation->get_Name());
    
    ASSERT_EQ(fieldUserName->GetFieldCode(), u" USERNAME ");
    ASSERT_EQ(fieldUserName->get_Result(), u"John Doe");
    
    // We can set this property to get our field to override the value currently stored in the UserInformation object.
    fieldUserName->set_UserName(u"Jane Doe");
    fieldUserName->Update();
    
    ASSERT_EQ(fieldUserName->GetFieldCode(), u" USERNAME  \"Jane Doe\"");
    ASSERT_EQ(fieldUserName->get_Result(), u"Jane Doe");
    
    // This does not affect the value in the UserInformation object.
    ASSERT_EQ(doc->get_FieldOptions()->get_CurrentUser()->get_Name(), u"John Doe");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.USERNAME.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.USERNAME.docx"));
    
    fieldUserName = System::ExplicitCast<Aspose::Words::Fields::FieldUserName>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldUserName, u" USERNAME  \"Jane Doe\"", u"Jane Doe", fieldUserName);
    ASSERT_EQ(fieldUserName->get_UserName(), u"Jane Doe");
}

namespace gtest_test
{

TEST_F(ExField, FieldUserName)
{
    s_instance->FieldUserName();
}

} // namespace gtest_test

void ExField::FieldStyleRefParagraphNumbers()
{
    //ExStart
    //ExFor:FieldStyleRef
    //ExFor:FieldStyleRef.InsertParagraphNumber
    //ExFor:FieldStyleRef.InsertParagraphNumberInFullContext
    //ExFor:FieldStyleRef.InsertParagraphNumberInRelativeContext
    //ExFor:FieldStyleRef.InsertRelativePosition
    //ExFor:FieldStyleRef.SearchFromBottom
    //ExFor:FieldStyleRef.StyleName
    //ExFor:FieldStyleRef.SuppressNonDelimiters
    //ExSummary:Shows how to use STYLEREF fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a list based using a Microsoft Word list template.
    System::SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(Lists::ListTemplate::NumberDefault);
    
    // This generated list will display "1.a )".
    // Space before the bracket is a non-delimiter character, which we can suppress.
    list->get_ListLevels()->idx_get(0)->set_NumberFormat(u"\x0000" u".");
    list->get_ListLevels()->idx_get(1)->set_NumberFormat(u"\x0001" u" )");
    
    // Add text and apply paragraph styles that STYLEREF fields will reference.
    builder->get_ListFormat()->set_List(list);
    builder->get_ListFormat()->ListIndent();
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"List Paragraph"));
    builder->Writeln(u"Item 1");
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Quote"));
    builder->Writeln(u"Item 2");
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"List Paragraph"));
    builder->Writeln(u"Item 3");
    builder->get_ListFormat()->RemoveNumbers();
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->idx_get(u"Normal"));
    
    // Place a STYLEREF field in the header and display the first "List Paragraph"-styled text in the document.
    builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
    field->set_StyleName(u"List Paragraph");
    
    // Place a STYLEREF field in the footer, and have it display the last text.
    builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
    field->set_StyleName(u"List Paragraph");
    field->set_SearchFromBottom(true);
    
    builder->MoveToDocumentEnd();
    
    // We can also use STYLEREF fields to reference the list numbers of lists.
    builder->Write(u"\nParagraph number: ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
    field->set_StyleName(u"Quote");
    field->set_InsertParagraphNumber(true);
    
    builder->Write(u"\nParagraph number, relative context: ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
    field->set_StyleName(u"Quote");
    field->set_InsertParagraphNumberInRelativeContext(true);
    
    builder->Write(u"\nParagraph number, full context: ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
    field->set_StyleName(u"Quote");
    field->set_InsertParagraphNumberInFullContext(true);
    
    builder->Write(u"\nParagraph number, full context, non-delimiter chars suppressed: ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(builder->InsertField(FieldType::FieldStyleRef, true));
    field->set_StyleName(u"Quote");
    field->set_InsertParagraphNumberInFullContext(true);
    field->set_SuppressNonDelimiters(true);
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.STYLEREF.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.STYLEREF.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  \"List Paragraph\"", u"Item 1", field);
    ASSERT_EQ(field->get_StyleName(), u"List Paragraph");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  \"List Paragraph\" \\l", u"Item 3", field);
    ASSERT_EQ(field->get_StyleName(), u"List Paragraph");
    ASSERT_TRUE(field->get_SearchFromBottom());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\n", u"‎b )", field);
    ASSERT_EQ(field->get_StyleName(), u"Quote");
    ASSERT_TRUE(field->get_InsertParagraphNumber());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\r", u"‎b )", field);
    ASSERT_EQ(field->get_StyleName(), u"Quote");
    ASSERT_TRUE(field->get_InsertParagraphNumberInRelativeContext());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(4));
    
    TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\w", u"‎1.b )", field);
    ASSERT_EQ(field->get_StyleName(), u"Quote");
    ASSERT_TRUE(field->get_InsertParagraphNumberInFullContext());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldStyleRef>(doc->get_Range()->get_Fields()->idx_get(5));
    
    TestUtil::VerifyField(FieldType::FieldStyleRef, u" STYLEREF  Quote \\w \\t", u"‎1.b)", field);
    ASSERT_EQ(field->get_StyleName(), u"Quote");
    ASSERT_TRUE(field->get_InsertParagraphNumberInFullContext());
    ASSERT_TRUE(field->get_SuppressNonDelimiters());
}

namespace gtest_test
{

TEST_F(ExField, FieldStyleRefParagraphNumbers)
{
    s_instance->FieldStyleRefParagraphNumbers();
}

} // namespace gtest_test

void ExField::FieldDate()
{
    //ExStart
    //ExFor:FieldDate
    //ExFor:FieldDate.UseLunarCalendar
    //ExFor:FieldDate.UseSakaEraCalendar
    //ExFor:FieldDate.UseUmAlQuraCalendar
    //ExFor:FieldDate.UseLastFormat
    //ExSummary:Shows how to use DATE fields to display dates according to different kinds of calendars.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // If we want the text in the document always to display the correct date, we can use a DATE field.
    // Below are three types of cultural calendars that a DATE field can use to display a date.
    // 1 -  Islamic Lunar Calendar:
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(builder->InsertField(FieldType::FieldDate, true));
    field->set_UseLunarCalendar(true);
    ASSERT_EQ(field->GetFieldCode(), u" DATE  \\h");
    builder->Writeln();
    
    // 2 -  Umm al-Qura calendar:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(builder->InsertField(FieldType::FieldDate, true));
    field->set_UseUmAlQuraCalendar(true);
    ASSERT_EQ(field->GetFieldCode(), u" DATE  \\u");
    builder->Writeln();
    
    // 3 -  Indian National Calendar:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(builder->InsertField(FieldType::FieldDate, true));
    field->set_UseSakaEraCalendar(true);
    ASSERT_EQ(field->GetFieldCode(), u" DATE  \\s");
    builder->Writeln();
    
    // Insert a DATE field and set its calendar type to the one last used by the host application.
    // In Microsoft Word, the type will be the most recently used in the Insert -> Text -> Date and Time dialog box.
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(builder->InsertField(FieldType::FieldDate, true));
    field->set_UseLastFormat(true);
    ASSERT_EQ(field->GetFieldCode(), u" DATE  \\l");
    builder->Writeln();
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.DATE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.DATE.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->get_Type(), FieldType::FieldDate);
    ASSERT_TRUE(field->get_UseLunarCalendar());
    ASSERT_EQ(field->GetFieldCode(), u" DATE  \\h");
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(doc->get_Range()->get_Fields()->idx_get(0)->get_Result(), u"\\d{1,2}[/]\\d{1,2}[/]\\d{4}")->get_Success());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\u", System::DateTime::get_Now().ToShortDateString(), field);
    ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\s", System::DateTime::get_Now().ToShortDateString(), field);
    ASSERT_TRUE(field->get_UseSakaEraCalendar());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldDate>(doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldDate, u" DATE  \\l", System::DateTime::get_Now().ToShortDateString(), field);
    ASSERT_TRUE(field->get_UseLastFormat());
}

namespace gtest_test
{

TEST_F(ExField, FieldDate)
{
    s_instance->FieldDate();
}

} // namespace gtest_test

void ExField::FieldCreateDate()
{
    //ExStart
    //ExFor:FieldCreateDate
    //ExFor:FieldCreateDate.UseLunarCalendar
    //ExFor:FieldCreateDate.UseSakaEraCalendar
    //ExFor:FieldCreateDate.UseUmAlQuraCalendar
    //ExSummary:Shows how to use the CREATEDATE field to display the creation date/time of the document.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Document.docx"));
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->Writeln(u" Date this document was created:");
    
    // We can use the CREATEDATE field to display the date and time of the creation of the document.
    // Below are three different calendar types according to which the CREATEDATE field can display the date/time.
    // 1 -  Islamic Lunar Calendar:
    builder->Write(u"According to the Lunar Calendar - ");
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
    field->set_UseLunarCalendar(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" CREATEDATE  \\h");
    
    // 2 -  Umm al-Qura calendar:
    builder->Write(u"\nAccording to the Umm al-Qura Calendar - ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
    field->set_UseUmAlQuraCalendar(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" CREATEDATE  \\u");
    
    // 3 -  Indian National Calendar:
    builder->Write(u"\nAccording to the Indian National Calendar - ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldCreateDate>(builder->InsertField(FieldType::FieldCreateDate, true));
    field->set_UseSakaEraCalendar(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" CREATEDATE  \\s");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.CREATEDATE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.CREATEDATE.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_CreatedTime(), System::DateTime(2017, 12, 5, 9, 56, 0));
    
    System::DateTime expectedDate = doc->get_BuiltInDocumentProperties()->get_CreatedTime().AddHours(System::TimeZoneInfo::get_Local()->GetUtcOffset(System::DateTime::get_UtcNow()).get_Hours());
    field = System::ExplicitCast<Aspose::Words::Fields::FieldCreateDate>(doc->get_Range()->get_Fields()->idx_get(0));
    System::SharedPtr<System::Globalization::Calendar> umAlQuraCalendar = System::MakeObject<System::Globalization::UmAlQuraCalendar>();
    
    TestUtil::VerifyField(FieldType::FieldCreateDate, u" CREATEDATE  \\h", System::String::Format(u"{0}/{1}/{2} ", umAlQuraCalendar->GetMonth(expectedDate), umAlQuraCalendar->GetDayOfMonth(expectedDate), umAlQuraCalendar->GetYear(expectedDate)) + expectedDate.AddHours(1).ToString(u"hh:mm:ss tt"), field);
    ASSERT_EQ(field->get_Type(), FieldType::FieldCreateDate);
    ASSERT_TRUE(field->get_UseLunarCalendar());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldCreateDate>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldCreateDate, u" CREATEDATE  \\u", System::String::Format(u"{0}/{1}/{2} ", umAlQuraCalendar->GetMonth(expectedDate), umAlQuraCalendar->GetDayOfMonth(expectedDate), umAlQuraCalendar->GetYear(expectedDate)) + expectedDate.AddHours(1).ToString(u"hh:mm:ss tt"), field);
    ASSERT_EQ(field->get_Type(), FieldType::FieldCreateDate);
    ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
}

namespace gtest_test
{

TEST_F(ExField, DISABLED_FieldCreateDate)
{
    s_instance->FieldCreateDate();
}

} // namespace gtest_test

void ExField::FieldSaveDate()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.LastSavedTime
    //ExFor:FieldSaveDate
    //ExFor:FieldSaveDate.UseLunarCalendar
    //ExFor:FieldSaveDate.UseSakaEraCalendar
    //ExFor:FieldSaveDate.UseUmAlQuraCalendar
    //ExSummary:Shows how to use the SAVEDATE field to display the date/time of the document's most recent save operation performed using Microsoft Word.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Document.docx"));
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->Writeln(u" Date this document was last saved:");
    
    // We can use the SAVEDATE field to display the last save operation's date and time on the document.
    // The save operation that these fields refer to is the manual save in an application like Microsoft Word,
    // not the document's Save method.
    // Below are three different calendar types according to which the SAVEDATE field can display the date/time.
    // 1 -  Islamic Lunar Calendar:
    builder->Write(u"According to the Lunar Calendar - ");
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
    field->set_UseLunarCalendar(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" SAVEDATE  \\h");
    
    // 2 -  Umm al-Qura calendar:
    builder->Write(u"\nAccording to the Umm al-Qura calendar - ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
    field->set_UseUmAlQuraCalendar(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" SAVEDATE  \\u");
    
    // 3 -  Indian National calendar:
    builder->Write(u"\nAccording to the Indian National calendar - ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSaveDate>(builder->InsertField(FieldType::FieldSaveDate, true));
    field->set_UseSakaEraCalendar(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" SAVEDATE  \\s");
    
    // The SAVEDATE fields draw their date/time values from the LastSavedTime built-in property.
    // The document's Save method will not update this value, but we can still update it manually.
    doc->get_BuiltInDocumentProperties()->set_LastSavedTime(System::DateTime::get_Now());
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.SAVEDATE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SAVEDATE.docx"));
    
    std::cout << doc->get_BuiltInDocumentProperties()->get_LastSavedTime() << std::endl;
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSaveDate>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->get_Type(), FieldType::FieldSaveDate);
    ASSERT_TRUE(field->get_UseLunarCalendar());
    ASSERT_EQ(field->GetFieldCode(), u" SAVEDATE  \\h");
    
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->get_Result(), u"\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M")->get_Success());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSaveDate>(doc->get_Range()->get_Fields()->idx_get(1));
    
    ASSERT_EQ(field->get_Type(), FieldType::FieldSaveDate);
    ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
    ASSERT_EQ(field->GetFieldCode(), u" SAVEDATE  \\u");
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->get_Result(), u"\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M")->get_Success());
}

namespace gtest_test
{

TEST_F(ExField, DISABLED_FieldSaveDate)
{
    s_instance->FieldSaveDate();
}

} // namespace gtest_test

void ExField::FieldBuilder()
{
    //ExStart
    //ExFor:FieldBuilder
    //ExFor:FieldBuilder.AddArgument(Int32)
    //ExFor:FieldBuilder.AddArgument(FieldArgumentBuilder)
    //ExFor:FieldBuilder.AddArgument(String)
    //ExFor:FieldBuilder.AddArgument(Double)
    //ExFor:FieldBuilder.AddArgument(FieldBuilder)
    //ExFor:FieldBuilder.AddSwitch(String)
    //ExFor:FieldBuilder.AddSwitch(String, Double)
    //ExFor:FieldBuilder.AddSwitch(String, Int32)
    //ExFor:FieldBuilder.AddSwitch(String, String)
    //ExFor:FieldBuilder.BuildAndInsert(Paragraph)
    //ExFor:FieldArgumentBuilder
    //ExFor:FieldArgumentBuilder.#ctor
    //ExFor:FieldArgumentBuilder.AddField(FieldBuilder)
    //ExFor:FieldArgumentBuilder.AddText(String)
    //ExFor:FieldArgumentBuilder.AddNode(Inline)
    //ExSummary:Shows how to construct fields using a field builder, and then insert them into the document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Below are three examples of field construction done using a field builder.
    // 1 -  Single field:
    // Use a field builder to add a SYMBOL field which displays the ƒ (Florin) symbol.
    auto builder = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldSymbol);
    builder->AddArgument(402);
    builder->AddSwitch(u"\\f", u"Arial");
    builder->AddSwitch(u"\\s", 25);
    builder->AddSwitch(u"\\u");
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph());
    
    ASSERT_EQ(field->GetFieldCode(), u" SYMBOL 402 \\f Arial \\s 25 \\u ");
    
    // 2 -  Nested field:
    // Use a field builder to create a formula field used as an inner field by another field builder.
    auto innerFormulaBuilder = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldFormula);
    innerFormulaBuilder->AddArgument(100);
    innerFormulaBuilder->AddArgument(u"+");
    innerFormulaBuilder->AddArgument(74);
    
    // Create another builder for another SYMBOL field, and insert the formula field
    // that we have created above into the SYMBOL field as its argument.
    builder = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldSymbol);
    builder->AddArgument(innerFormulaBuilder);
    field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->AppendParagraph(System::String::Empty));
    
    // The outer SYMBOL field will use the formula field result, 174, as its argument,
    // which will make the field display the ® (Registered Sign) symbol since its character number is 174.
    ASSERT_EQ(field->GetFieldCode(), u" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ");
    
    // 3 -  Multiple nested fields and arguments:
    // Now, we will use a builder to create an IF field, which displays one of two custom string values,
    // depending on the true/false value of its expression. To get a true/false value
    // that determines which string the IF field displays, the IF field will test two numeric expressions for equality.
    // We will provide the two expressions in the form of formula fields, which we will nest inside the IF field.
    auto leftExpression = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldFormula);
    leftExpression->AddArgument(2);
    leftExpression->AddArgument(u"+");
    leftExpression->AddArgument(3);
    
    auto rightExpression = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldFormula);
    rightExpression->AddArgument(2.5);
    rightExpression->AddArgument(u"*");
    rightExpression->AddArgument(5.2);
    
    // Next, we will build two field arguments, which will serve as the true/false output strings for the IF field.
    // These arguments will reuse the output values of our numeric expressions.
    auto trueOutput = System::MakeObject<Aspose::Words::Fields::FieldArgumentBuilder>();
    trueOutput->AddText(u"True, both expressions amount to ");
    trueOutput->AddField(leftExpression);
    
    auto falseOutput = System::MakeObject<Aspose::Words::Fields::FieldArgumentBuilder>();
    falseOutput->AddNode(System::MakeObject<Aspose::Words::Run>(doc, u"False, "));
    falseOutput->AddField(leftExpression);
    falseOutput->AddNode(System::MakeObject<Aspose::Words::Run>(doc, u" does not equal "));
    falseOutput->AddField(rightExpression);
    
    // Finally, we will create one more field builder for the IF field and combine all of the expressions.
    builder = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIf);
    builder->AddArgument(leftExpression);
    builder->AddArgument(u"=");
    builder->AddArgument(rightExpression);
    builder->AddArgument(trueOutput);
    builder->AddArgument(falseOutput);
    field = builder->BuildAndInsert(doc->get_FirstSection()->get_Body()->AppendParagraph(System::String::Empty));
    
    ASSERT_EQ(field->GetFieldCode(), (System::String(u" IF \u0013 = 2 + 3 \u0014\u0015 = \u0013 = 2.5 * 5.2 \u0014\u0015 ") + u"\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " + u"\"False, \u0013 = 2 + 3 \u0014\u0015 does not equal \u0013 = 2.5 * 5.2 \u0014\u0015\" "));
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.SYMBOL.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SYMBOL.docx"));
    
    auto fieldSymbol = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL 402 \\f Arial \\s 25 \\u ", System::String::Empty, fieldSymbol);
    ASSERT_EQ(fieldSymbol->get_DisplayResult(), u"ƒ");
    
    fieldSymbol = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", System::String::Empty, fieldSymbol);
    ASSERT_EQ(fieldSymbol->get_DisplayResult(), u"®");
    
    TestUtil::VerifyField(FieldType::FieldFormula, u" = 100 + 74 ", u"174", doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldIf, System::String(u" IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 ") + u"\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " + u"\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ", u"False, 5 does not equal 13", doc->get_Range()->get_Fields()->idx_get(3));
    TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 + 3 ", u"5", doc->get_Range()->get_Fields()->idx_get(4));
    TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(4), doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldFormula, u" = 2.5 * 5.2 ", u"13", doc->get_Range()->get_Fields()->idx_get(5));
    TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(5), doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 + 3 ", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(6));
    TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(6), doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 + 3 ", u"5", doc->get_Range()->get_Fields()->idx_get(7));
    TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(7), doc->get_Range()->get_Fields()->idx_get(3));
    
    TestUtil::VerifyField(FieldType::FieldFormula, u" = 2.5 * 5.2 ", u"13", doc->get_Range()->get_Fields()->idx_get(8));
    TestUtil::FieldsAreNested(doc->get_Range()->get_Fields()->idx_get(8), doc->get_Range()->get_Fields()->idx_get(3));
}

namespace gtest_test
{

TEST_F(ExField, FieldBuilder)
{
    s_instance->FieldBuilder();
}

} // namespace gtest_test

void ExField::FieldAuthor()
{
    //ExStart
    //ExFor:FieldAuthor
    //ExFor:FieldAuthor.AuthorName
    //ExFor:FieldOptions.DefaultDocumentAuthor
    //ExSummary:Shows how to use an AUTHOR field to display a document creator's name.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // AUTHOR fields source their results from the built-in document property called "Author".
    // If we create and save a document in Microsoft Word,
    // it will have our username in that property.
    // However, if we create a document programmatically using Aspose.Words,
    // the "Author" property, by default, will be an empty string.
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Author(), System::String::Empty);
    
    // Set a backup author name for AUTHOR fields to use
    // if the "Author" property contains an empty string.
    doc->get_FieldOptions()->set_DefaultDocumentAuthor(u"Joe Bloggs");
    
    builder->Write(u"This document was created by ");
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAuthor>(builder->InsertField(FieldType::FieldAuthor, true));
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" AUTHOR ");
    ASSERT_EQ(field->get_Result(), u"Joe Bloggs");
    
    // Updating an AUTHOR field that contains a value
    // will apply that value to the "Author" built-in property.
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Author(), u"Joe Bloggs");
    
    // Changing this property, then updating the AUTHOR field will apply this value to the field.
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" AUTHOR ");
    ASSERT_EQ(field->get_Result(), u"John Doe");
    
    // If we update an AUTHOR field after changing its "Name" property,
    // then the field will display the new name and apply the new name to the built-in property.
    field->set_AuthorName(u"Jane Doe");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" AUTHOR  \"Jane Doe\"");
    ASSERT_EQ(field->get_Result(), u"Jane Doe");
    
    // AUTHOR fields do not affect the DefaultDocumentAuthor property.
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Author(), u"Jane Doe");
    ASSERT_EQ(doc->get_FieldOptions()->get_DefaultDocumentAuthor(), u"Joe Bloggs");
    
    doc->Save(get_ArtifactsDir() + u"Field.AUTHOR.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.AUTHOR.docx"));
    
    ASSERT_TRUE(System::TestTools::IsNull(doc->get_FieldOptions()->get_DefaultDocumentAuthor()));
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Author(), u"Jane Doe");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldAuthor>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldAuthor, u" AUTHOR  \"Jane Doe\"", u"Jane Doe", field);
    ASSERT_EQ(field->get_AuthorName(), u"Jane Doe");
}

namespace gtest_test
{

TEST_F(ExField, FieldAuthor)
{
    s_instance->FieldAuthor();
}

} // namespace gtest_test

void ExField::FieldDocVariable()
{
    //ExStart
    //ExFor:FieldDocProperty
    //ExFor:FieldDocVariable
    //ExFor:FieldDocVariable.VariableName
    //ExSummary:Shows how to use DOCPROPERTY fields to display document properties and variables.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are two ways of using DOCPROPERTY fields.
    // 1 -  Display a built-in property:
    // Set a custom value for the "Category" built-in property, then insert a DOCPROPERTY field that references it.
    doc->get_BuiltInDocumentProperties()->set_Category(u"My category");
    
    auto fieldDocProperty = System::ExplicitCast<Aspose::Words::Fields::FieldDocProperty>(builder->InsertField(u" DOCPROPERTY Category "));
    fieldDocProperty->Update();
    
    ASSERT_EQ(fieldDocProperty->GetFieldCode(), u" DOCPROPERTY Category ");
    ASSERT_EQ(fieldDocProperty->get_Result(), u"My category");
    
    builder->InsertParagraph();
    
    // 2 -  Display a custom document variable:
    // Define a custom variable, then reference that variable with a DOCPROPERTY field.
    ASSERT_EQ(doc->get_Variables()->get_Count(), 0);
    doc->get_Variables()->Add(u"My variable", u"My variable's value");
    
    auto fieldDocVariable = System::ExplicitCast<Aspose::Words::Fields::FieldDocVariable>(builder->InsertField(FieldType::FieldDocVariable, true));
    fieldDocVariable->set_VariableName(u"My Variable");
    fieldDocVariable->Update();
    
    ASSERT_EQ(fieldDocVariable->GetFieldCode(), u" DOCVARIABLE  \"My Variable\"");
    ASSERT_EQ(fieldDocVariable->get_Result(), u"My variable's value");
    
    doc->Save(get_ArtifactsDir() + u"Field.DOCPROPERTY.DOCVARIABLE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.DOCPROPERTY.DOCVARIABLE.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Category(), u"My category");
    
    fieldDocProperty = System::ExplicitCast<Aspose::Words::Fields::FieldDocProperty>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldDocProperty, u" DOCPROPERTY Category ", u"My category", fieldDocProperty);
    
    fieldDocVariable = System::ExplicitCast<Aspose::Words::Fields::FieldDocVariable>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldDocVariable, u" DOCVARIABLE  \"My Variable\"", u"My variable's value", fieldDocVariable);
    ASSERT_EQ(fieldDocVariable->get_VariableName(), u"My Variable");
}

namespace gtest_test
{

TEST_F(ExField, FieldDocVariable)
{
    s_instance->FieldDocVariable();
}

} // namespace gtest_test

void ExField::FieldSubject()
{
    //ExStart
    //ExFor:FieldSubject
    //ExFor:FieldSubject.Text
    //ExSummary:Shows how to use the SUBJECT field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Set a value for the document's "Subject" built-in property.
    doc->get_BuiltInDocumentProperties()->set_Subject(u"My subject");
    
    // Create a SUBJECT field to display the value of that built-in property.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldSubject>(builder->InsertField(FieldType::FieldSubject, true));
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" SUBJECT ");
    ASSERT_EQ(field->get_Result(), u"My subject");
    
    // If we give the SUBJECT field's Text property value and update it, the field will
    // overwrite the current value of the "Subject" built-in property with the value of its Text property,
    // and then display the new value.
    field->set_Text(u"My new subject");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" SUBJECT  \"My new subject\"");
    ASSERT_EQ(field->get_Result(), u"My new subject");
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Subject(), u"My new subject");
    
    doc->Save(get_ArtifactsDir() + u"Field.SUBJECT.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SUBJECT.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Subject(), u"My new subject");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSubject>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldSubject, u" SUBJECT  \"My new subject\"", u"My new subject", field);
    ASSERT_EQ(field->get_Text(), u"My new subject");
}

namespace gtest_test
{

TEST_F(ExField, FieldSubject)
{
    s_instance->FieldSubject();
}

} // namespace gtest_test

void ExField::FieldComments()
{
    //ExStart
    //ExFor:FieldComments
    //ExFor:FieldComments.Text
    //ExSummary:Shows how to use the COMMENTS field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set a value for the document's "Comments" built-in property.
    doc->get_BuiltInDocumentProperties()->set_Comments(u"My comment.");
    
    // Create a COMMENTS field to display the value of that built-in property.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldComments>(builder->InsertField(FieldType::FieldComments, true));
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" COMMENTS ");
    ASSERT_EQ(field->get_Result(), u"My comment.");
    
    // If we give the COMMENTS field's Text property value and update it, the field will
    // overwrite the current value of the "Comments" built-in property with the value of its Text property,
    // and then display the new value.
    field->set_Text(u"My overriding comment.");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" COMMENTS  \"My overriding comment.\"");
    ASSERT_EQ(field->get_Result(), u"My overriding comment.");
    
    doc->Save(get_ArtifactsDir() + u"Field.COMMENTS.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.COMMENTS.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Comments(), u"My overriding comment.");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldComments>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldComments, u" COMMENTS  \"My overriding comment.\"", u"My overriding comment.", field);
    ASSERT_EQ(field->get_Text(), u"My overriding comment.");
}

namespace gtest_test
{

TEST_F(ExField, FieldComments)
{
    s_instance->FieldComments();
}

} // namespace gtest_test

void ExField::FieldFileSize()
{
    //ExStart
    //ExFor:FieldFileSize
    //ExFor:FieldFileSize.IsInKilobytes
    //ExFor:FieldFileSize.IsInMegabytes
    //ExSummary:Shows how to display the file size of a document with a FILESIZE field.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Document.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Bytes(), 18105);
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->InsertParagraph();
    
    // Below are three different units of measure
    // with which FILESIZE fields can display the document's file size.
    // 1 -  Bytes:
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" FILESIZE ");
    ASSERT_EQ(field->get_Result(), u"18105");
    
    // 2 -  Kilobytes:
    builder->InsertParagraph();
    field = System::ExplicitCast<Aspose::Words::Fields::FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
    field->set_IsInKilobytes(true);
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" FILESIZE  \\k");
    ASSERT_EQ(field->get_Result(), u"18");
    
    // 3 -  Megabytes:
    builder->InsertParagraph();
    field = System::ExplicitCast<Aspose::Words::Fields::FieldFileSize>(builder->InsertField(FieldType::FieldFileSize, true));
    field->set_IsInMegabytes(true);
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" FILESIZE  \\m");
    ASSERT_EQ(field->get_Result(), u"0");
    
    // To update the values of these fields while editing in Microsoft Word,
    // we must first save the changes, and then manually update these fields.
    doc->Save(get_ArtifactsDir() + u"Field.FILESIZE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.FILESIZE.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldFileSize>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldFileSize, u" FILESIZE ", u"18105", field);
    
    // These fields will need to be updated to produce an accurate result.
    doc->UpdateFields();
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldFileSize>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldFileSize, u" FILESIZE  \\k", u"13", field);
    ASSERT_TRUE(field->get_IsInKilobytes());
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldFileSize>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldFileSize, u" FILESIZE  \\m", u"0", field);
    ASSERT_TRUE(field->get_IsInMegabytes());
}

namespace gtest_test
{

TEST_F(ExField, FieldFileSize)
{
    s_instance->FieldFileSize();
}

} // namespace gtest_test

void ExField::FieldGoToButton()
{
    //ExStart
    //ExFor:FieldGoToButton
    //ExFor:FieldGoToButton.DisplayText
    //ExFor:FieldGoToButton.Location
    //ExSummary:Shows to insert a GOTOBUTTON field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add a GOTOBUTTON field. When we double-click this field in Microsoft Word,
    // it will take the text cursor to the bookmark whose name the Location property references.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldGoToButton>(builder->InsertField(FieldType::FieldGoToButton, true));
    field->set_DisplayText(u"My Button");
    field->set_Location(u"MyBookmark");
    
    ASSERT_EQ(field->GetFieldCode(), u" GOTOBUTTON  MyBookmark My Button");
    
    // Insert a valid bookmark for the field to reference.
    builder->InsertBreak(BreakType::PageBreak);
    builder->StartBookmark(field->get_Location());
    builder->Writeln(u"Bookmark text contents.");
    builder->EndBookmark(field->get_Location());
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.GOTOBUTTON.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.GOTOBUTTON.docx"));
    field = System::ExplicitCast<Aspose::Words::Fields::FieldGoToButton>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldGoToButton, u" GOTOBUTTON  MyBookmark My Button", System::String::Empty, field);
    ASSERT_EQ(field->get_DisplayText(), u"My Button");
    ASSERT_EQ(field->get_Location(), u"MyBookmark");
}

namespace gtest_test
{

TEST_F(ExField, FieldGoToButton)
{
    s_instance->FieldGoToButton();
}

} // namespace gtest_test

void ExField::FieldFillIn()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a FILLIN field. When we manually update this field in Microsoft Word,
    // it will prompt us to enter a response. The field will then display the response as text.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldFillIn>(builder->InsertField(FieldType::FieldFillIn, true));
    field->set_PromptText(u"Please enter a response:");
    field->set_DefaultResponse(u"A default response.");
    
    // We can also use these fields to ask the user for a unique response for each page
    // created during a mail merge done using Microsoft Word.
    field->set_PromptOnceOnMailMerge(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o");
    
    auto mergeField = System::ExplicitCast<Aspose::Words::Fields::FieldMergeField>(builder->InsertField(FieldType::FieldMergeField, true));
    mergeField->set_FieldName(u"MergeField");
    
    // If we perform a mail merge programmatically, we can use a custom prompt respondent
    // to automatically edit responses for FILLIN fields that the mail merge encounters.
    doc->get_FieldOptions()->set_UserPromptRespondent(System::MakeObject<Aspose::Words::ApiExamples::ExField::PromptRespondent>());
    doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"MergeField"}), System::MakeArray<System::SharedPtr<System::Object>>({System::ExplicitCast<System::Object>(u"")}));
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.FILLIN.docx");
    TestFieldFillIn(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.FILLIN.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldFillIn)
{
    s_instance->FieldFillIn();
}

} // namespace gtest_test

void ExField::FieldInfo()
{
    //ExStart
    //ExFor:FieldInfo
    //ExFor:FieldInfo.InfoType
    //ExFor:FieldInfo.NewValue
    //ExSummary:Shows how to work with INFO fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Set a value for the "Comments" built-in property and then insert an INFO field to display that property's value.
    doc->get_BuiltInDocumentProperties()->set_Comments(u"My comment");
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldInfo>(builder->InsertField(FieldType::FieldInfo, true));
    field->set_InfoType(u"Comments");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" INFO  Comments");
    ASSERT_EQ(field->get_Result(), u"My comment");
    
    builder->Writeln();
    
    // Setting a value for the field's NewValue property and updating
    // the field will also overwrite the corresponding built-in property with the new value.
    field = System::ExplicitCast<Aspose::Words::Fields::FieldInfo>(builder->InsertField(FieldType::FieldInfo, true));
    field->set_InfoType(u"Comments");
    field->set_NewValue(u"New comment");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" INFO  Comments \"New comment\"");
    ASSERT_EQ(field->get_Result(), u"New comment");
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Comments(), u"New comment");
    
    doc->Save(get_ArtifactsDir() + u"Field.INFO.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.INFO.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Comments(), u"New comment");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldInfo>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldInfo, u" INFO  Comments", u"My comment", field);
    ASSERT_EQ(field->get_InfoType(), u"Comments");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldInfo>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldInfo, u" INFO  Comments \"New comment\"", u"New comment", field);
    ASSERT_EQ(field->get_InfoType(), u"Comments");
    ASSERT_EQ(field->get_NewValue(), u"New comment");
}

namespace gtest_test
{

TEST_F(ExField, FieldInfo)
{
    s_instance->FieldInfo();
}

} // namespace gtest_test

void ExField::FieldMacroButton()
{
    //ExStart
    //ExFor:Document.HasMacros
    //ExFor:FieldMacroButton
    //ExFor:FieldMacroButton.DisplayText
    //ExFor:FieldMacroButton.MacroName
    //ExSummary:Shows how to use MACROBUTTON fields to allow us to run a document's macros by clicking.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Macro.docm"));
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    ASSERT_TRUE(doc->get_HasMacros());
    
    // Insert a MACROBUTTON field, and reference one of the document's macros by name in the MacroName property.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldMacroButton>(builder->InsertField(FieldType::FieldMacroButton, true));
    field->set_MacroName(u"MyMacro");
    field->set_DisplayText(System::String(u"Double click to run macro: ") + field->get_MacroName());
    
    ASSERT_EQ(field->GetFieldCode(), u" MACROBUTTON  MyMacro Double click to run macro: MyMacro");
    
    // Use the property to reference "ViewZoom200", a macro that ships with Microsoft Word.
    // We can find all other macros via View -> Macros (dropdown) -> View Macros.
    // In that menu, select "Word Commands" from the "Macros in:" drop down.
    // If our document contains a custom macro with the same name as a stock macro,
    // our macro will be the one that the MACROBUTTON field runs.
    builder->InsertParagraph();
    field = System::ExplicitCast<Aspose::Words::Fields::FieldMacroButton>(builder->InsertField(FieldType::FieldMacroButton, true));
    field->set_MacroName(u"ViewZoom200");
    field->set_DisplayText(System::String(u"Run ") + field->get_MacroName());
    
    ASSERT_EQ(field->GetFieldCode(), u" MACROBUTTON  ViewZoom200 Run ViewZoom200");
    
    // Save the document as a macro-enabled document type.
    doc->Save(get_ArtifactsDir() + u"Field.MACROBUTTON.docm");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.MACROBUTTON.docm"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldMacroButton>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldMacroButton, u" MACROBUTTON  MyMacro Double click to run macro: MyMacro", System::String::Empty, field);
    ASSERT_EQ(field->get_MacroName(), u"MyMacro");
    ASSERT_EQ(field->get_DisplayText(), u"Double click to run macro: MyMacro");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldMacroButton>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldMacroButton, u" MACROBUTTON  ViewZoom200 Run ViewZoom200", System::String::Empty, field);
    ASSERT_EQ(field->get_MacroName(), u"ViewZoom200");
    ASSERT_EQ(field->get_DisplayText(), u"Run ViewZoom200");
}

namespace gtest_test
{

TEST_F(ExField, FieldMacroButton)
{
    s_instance->FieldMacroButton();
}

} // namespace gtest_test

void ExField::FieldKeywords()
{
    //ExStart
    //ExFor:FieldKeywords
    //ExFor:FieldKeywords.Text
    //ExSummary:Shows to insert a KEYWORDS field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add some keywords, also referred to as "tags" in File Explorer.
    doc->get_BuiltInDocumentProperties()->set_Keywords(u"Keyword1, Keyword2");
    
    // The KEYWORDS field displays the value of this property.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldKeywords>(builder->InsertField(FieldType::FieldKeyword, true));
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" KEYWORDS ");
    ASSERT_EQ(field->get_Result(), (u"Keyword1, Keyword2"));
    
    // Setting a value for the field's Text property,
    // and then updating the field will also overwrite the corresponding built-in property with the new value.
    field->set_Text(u"OverridingKeyword");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" KEYWORDS  OverridingKeyword");
    ASSERT_EQ(field->get_Result(), u"OverridingKeyword");
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Keywords(), u"OverridingKeyword");
    
    doc->Save(get_ArtifactsDir() + u"Field.KEYWORDS.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.KEYWORDS.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Keywords(), u"OverridingKeyword");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldKeywords>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldKeyword, u" KEYWORDS  OverridingKeyword", u"OverridingKeyword", field);
    ASSERT_EQ(field->get_Text(), u"OverridingKeyword");
}

namespace gtest_test
{

TEST_F(ExField, FieldKeywords)
{
    s_instance->FieldKeywords();
}

} // namespace gtest_test

void ExField::FieldNum()
{
    //ExStart
    //ExFor:FieldPage
    //ExFor:FieldNumChars
    //ExFor:FieldNumPages
    //ExFor:FieldNumWords
    //ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Paragraphs.docx"));
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
    builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
    
    // Below are three types of fields that we can use to track the size of our documents.
    // 1 -  Track the character count with a NUMCHARS field:
    auto fieldNumChars = System::ExplicitCast<Aspose::Words::Fields::FieldNumChars>(builder->InsertField(FieldType::FieldNumChars, true));
    builder->Writeln(u" characters");
    
    // 2 -  Track the word count with a NUMWORDS field:
    auto fieldNumWords = System::ExplicitCast<Aspose::Words::Fields::FieldNumWords>(builder->InsertField(FieldType::FieldNumWords, true));
    builder->Writeln(u" words");
    
    // 3 -  Use both PAGE and NUMPAGES fields to display what page the field is on,
    // and the total number of pages in the document:
    builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
    builder->Write(u"Page ");
    auto fieldPage = System::ExplicitCast<Aspose::Words::Fields::FieldPage>(builder->InsertField(FieldType::FieldPage, true));
    builder->Write(u" of ");
    auto fieldNumPages = System::ExplicitCast<Aspose::Words::Fields::FieldNumPages>(builder->InsertField(FieldType::FieldNumPages, true));
    
    ASSERT_EQ(fieldNumChars->GetFieldCode(), u" NUMCHARS ");
    ASSERT_EQ(fieldNumWords->GetFieldCode(), u" NUMWORDS ");
    ASSERT_EQ(fieldNumPages->GetFieldCode(), u" NUMPAGES ");
    ASSERT_EQ(fieldPage->GetFieldCode(), u" PAGE ");
    
    // These fields will not maintain accurate values in real time
    // while we edit the document programmatically using Aspose.Words, or in Microsoft Word.
    // We need to update them every we need to see an up-to-date value.
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx"));
    
    TestUtil::VerifyField(FieldType::FieldNumChars, u" NUMCHARS ", u"6009", doc->get_Range()->get_Fields()->idx_get(0));
    TestUtil::VerifyField(FieldType::FieldNumWords, u" NUMWORDS ", u"1054", doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldPage, u" PAGE ", u"6", doc->get_Range()->get_Fields()->idx_get(2));
    TestUtil::VerifyField(FieldType::FieldNumPages, u" NUMPAGES ", u"6", doc->get_Range()->get_Fields()->idx_get(3));
}

namespace gtest_test
{

TEST_F(ExField, FieldNum)
{
    s_instance->FieldNum();
}

} // namespace gtest_test

void ExField::FieldPrint()
{
    //ExStart
    //ExFor:FieldPrint
    //ExFor:FieldPrint.PostScriptGroup
    //ExFor:FieldPrint.PrinterInstructions
    //ExSummary:Shows to insert a PRINT field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"My paragraph");
    
    // The PRINT field can send instructions to the printer.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldPrint>(builder->InsertField(FieldType::FieldPrint, true));
    
    // Set the area for the printer to perform instructions over.
    // In this case, it will be the paragraph that contains our PRINT field.
    field->set_PostScriptGroup(u"para");
    
    // When we use a printer that supports PostScript to print our document,
    // this command will turn the entire area that we specified in "field.PostScriptGroup" white.
    field->set_PrinterInstructions(u"erasepage");
    
    ASSERT_EQ(field->GetFieldCode(), u" PRINT  erasepage \\p para");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.PRINT.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.PRINT.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldPrint>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldPrint, u" PRINT  erasepage \\p para", System::String::Empty, field);
    ASSERT_EQ(field->get_PostScriptGroup(), u"para");
    ASSERT_EQ(field->get_PrinterInstructions(), u"erasepage");
}

namespace gtest_test
{

TEST_F(ExField, FieldPrint)
{
    s_instance->FieldPrint();
}

} // namespace gtest_test

void ExField::FieldPrintDate()
{
    //ExStart
    //ExFor:FieldPrintDate
    //ExFor:FieldPrintDate.UseLunarCalendar
    //ExFor:FieldPrintDate.UseSakaEraCalendar
    //ExFor:FieldPrintDate.UseUmAlQuraCalendar
    //ExSummary:Shows read PRINTDATE fields.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Field sample - PRINTDATE.docx"));
    
    // When a document is printed by a printer or printed as a PDF (but not exported to PDF),
    // PRINTDATE fields will display the print operation's date/time.
    // If no printing has taken place, these fields will display "0/0/0000".
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->get_Result(), u"3/25/2020 12:00:00 AM");
    ASSERT_EQ(field->GetFieldCode(), u" PRINTDATE ");
    
    // Below are three different calendar types according to which the PRINTDATE field
    // can display the date and time of the last printing operation.
    // 1 -  Islamic Lunar Calendar:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(1));
    
    ASSERT_TRUE(field->get_UseLunarCalendar());
    ASSERT_EQ(field->get_Result(), u"8/1/1441 12:00:00 AM");
    ASSERT_EQ(field->GetFieldCode(), u" PRINTDATE  \\h");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(2));
    
    // 2 -  Umm al-Qura calendar:
    ASSERT_TRUE(field->get_UseUmAlQuraCalendar());
    ASSERT_EQ(field->get_Result(), u"8/1/1441 12:00:00 AM");
    ASSERT_EQ(field->GetFieldCode(), u" PRINTDATE  \\u");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldPrintDate>(doc->get_Range()->get_Fields()->idx_get(3));
    
    // 3 -  Indian National Calendar:
    ASSERT_TRUE(field->get_UseSakaEraCalendar());
    ASSERT_EQ(field->get_Result(), u"1/5/1942 12:00:00 AM");
    ASSERT_EQ(field->GetFieldCode(), u" PRINTDATE  \\s");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, FieldPrintDate)
{
    s_instance->FieldPrintDate();
}

} // namespace gtest_test

void ExField::FieldQuote()
{
    //ExStart
    //ExFor:FieldQuote
    //ExFor:FieldQuote.Text
    //ExFor:Document.UpdateFields
    //ExSummary:Shows to use the QUOTE field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a QUOTE field, which will display the value of its Text property.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldQuote>(builder->InsertField(FieldType::FieldQuote, true));
    field->set_Text(u"\"Quoted text\"");
    
    ASSERT_EQ(field->GetFieldCode(), u" QUOTE  \"\\\"Quoted text\\\"\"");
    
    // Insert a QUOTE field and nest a DATE field inside it.
    // DATE fields update their value to the current date every time we open the document using Microsoft Word.
    // Nesting the DATE field inside the QUOTE field like this will freeze its value
    // to the date when we created the document.
    builder->Write(u"\nDocument creation date: ");
    field = System::ExplicitCast<Aspose::Words::Fields::FieldQuote>(builder->InsertField(FieldType::FieldQuote, true));
    builder->MoveTo(field->get_Separator());
    builder->InsertField(FieldType::FieldDate, true);
    
    ASSERT_EQ(field->GetFieldCode(), System::String(u" QUOTE \u0013 DATE \u0014") + System::DateTime::get_Now().get_Date().ToShortDateString() + u"\u0015");
    
    // Update all the fields to display their correct results.
    doc->UpdateFields();
    
    ASSERT_EQ(doc->get_Range()->get_Fields()->idx_get(0)->get_Result(), u"\"Quoted text\"");
    
    doc->Save(get_ArtifactsDir() + u"Field.QUOTE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.QUOTE.docx"));
    
    TestUtil::VerifyField(FieldType::FieldQuote, u" QUOTE  \"\\\"Quoted text\\\"\"", u"\"Quoted text\"", doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldQuote, System::String(u" QUOTE \u0013 DATE \u0014") + System::DateTime::get_Now().get_Date().ToShortDateString() + u"\u0015", System::DateTime::get_Now().get_Date().ToShortDateString(), doc->get_Range()->get_Fields()->idx_get(1));
}

namespace gtest_test
{

TEST_F(ExField, FieldQuote)
{
    s_instance->FieldQuote();
}

} // namespace gtest_test

void ExField::FieldNoteRef()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a bookmark with a footnote that the NOTEREF field will reference.
    InsertBookmarkWithFootnote(builder, u"MyBookmark1", u"Contents of MyBookmark1", u"Footnote from MyBookmark1");
    
    // This NOTEREF field will display the number of the footnote inside the referenced bookmark.
    // Setting the InsertHyperlink property lets us jump to the bookmark by Ctrl + clicking the field in Microsoft Word.
    ASSERT_EQ(InsertFieldNoteRef(builder, u"MyBookmark2", true, false, false, u"Hyperlink to Bookmark2, with footnote number ")->GetFieldCode(), u" NOTEREF  MyBookmark2 \\h");
    
    // When using the \p flag, after the footnote number, the field also displays the bookmark's position relative to the field.
    // Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update.
    ASSERT_EQ(InsertFieldNoteRef(builder, u"MyBookmark1", true, true, false, u"Bookmark1, with footnote number ")->GetFieldCode(), u" NOTEREF  MyBookmark1 \\h \\p");
    
    // Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below".
    // The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text.
    ASSERT_EQ(InsertFieldNoteRef(builder, u"MyBookmark2", true, true, true, u"Bookmark2, with footnote number ")->GetFieldCode(), u" NOTEREF  MyBookmark2 \\h \\p \\f");
    
    builder->InsertBreak(BreakType::PageBreak);
    InsertBookmarkWithFootnote(builder, u"MyBookmark2", u"Contents of MyBookmark2", u"Footnote from MyBookmark2");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.NOTEREF.docx");
    TestNoteRef(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.NOTEREF.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldNoteRef)
{
    s_instance->FieldNoteRef();
}

} // namespace gtest_test

void ExField::NoteRef()
{
    //ExStart
    //ExFor:FieldNoteRef
    //ExSummary:Shows how to cross-reference footnotes with the NOTEREF field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"CrossReference: ");
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldNoteRef>(builder->InsertField(FieldType::FieldNoteRef, false));
    // <--- don't update field
    field->set_BookmarkName(u"CrossRefBookmark");
    field->set_InsertHyperlink(true);
    field->set_InsertReferenceMark(true);
    field->set_InsertRelativePosition(false);
    builder->Writeln();
    
    builder->StartBookmark(u"CrossRefBookmark");
    builder->Write(u"Hello world!");
    builder->InsertFootnote(FootnoteType::Footnote, u"Cross referenced footnote.");
    builder->EndBookmark(u"CrossRefBookmark");
    builder->Writeln();
    
    doc->UpdateFields();
    
    // This field works only in older versions of Microsoft Word.
    doc->Save(get_ArtifactsDir() + u"Field.NOTEREF.doc");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.NOTEREF.doc"));
    field = System::ExplicitCast<Aspose::Words::Fields::FieldNoteRef>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldNoteRef, u" NOTEREF  CrossRefBookmark \\h \\f", u"1", field);
    TestUtil::VerifyFootnote(FootnoteType::Footnote, true, nullptr, u"Cross referenced footnote.", System::ExplicitCast<Aspose::Words::Notes::Footnote>(doc->GetChild(NodeType::Footnote, 0, true)));
}

namespace gtest_test
{

TEST_F(ExField, NoteRef)
{
    s_instance->NoteRef();
}

} // namespace gtest_test

void ExField::FieldPageRef()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    InsertAndNameBookmark(builder, u"MyBookmark1");
    
    // Insert a PAGEREF field that displays what page a bookmark is on.
    // Set the InsertHyperlink flag to make the field also function as a clickable link to the bookmark.
    ASSERT_EQ(InsertFieldPageRef(builder, u"MyBookmark3", true, false, u"Hyperlink to Bookmark3, on page: ")->GetFieldCode(), u" PAGEREF  MyBookmark3 \\h");
    
    // We can use the \p flag to get the PAGEREF field to display
    // the bookmark's position relative to the position of the field.
    // Bookmark1 is on the same page and above this field, so this field's displayed result will be "above".
    ASSERT_EQ(InsertFieldPageRef(builder, u"MyBookmark1", true, true, u"Bookmark1 is ")->GetFieldCode(), u" PAGEREF  MyBookmark1 \\h \\p");
    
    // Bookmark2 will be on the same page and below this field, so this field's displayed result will be "below".
    ASSERT_EQ(InsertFieldPageRef(builder, u"MyBookmark2", true, true, u"Bookmark2 is ")->GetFieldCode(), u" PAGEREF  MyBookmark2 \\h \\p");
    
    // Bookmark3 will be on a different page, so the field will display "on page 2".
    ASSERT_EQ(InsertFieldPageRef(builder, u"MyBookmark3", true, true, u"Bookmark3 is ")->GetFieldCode(), u" PAGEREF  MyBookmark3 \\h \\p");
    
    InsertAndNameBookmark(builder, u"MyBookmark2");
    builder->InsertBreak(BreakType::PageBreak);
    InsertAndNameBookmark(builder, u"MyBookmark3");
    
    doc->UpdatePageLayout();
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.PAGEREF.docx");
    TestPageRef(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.PAGEREF.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldPageRef)
{
    s_instance->FieldPageRef();
}

} // namespace gtest_test

void ExField::FieldRef()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartBookmark(u"MyBookmark");
    builder->InsertFootnote(FootnoteType::Footnote, u"MyBookmark footnote #1");
    builder->Write(u"Text that will appear in REF field");
    builder->InsertFootnote(FootnoteType::Footnote, u"MyBookmark footnote #2");
    builder->EndBookmark(u"MyBookmark");
    builder->MoveToDocumentStart();
    
    // We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at.
    builder->get_ListFormat()->ApplyNumberDefault();
    builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u"> \x0000");
    
    // Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes.
    System::SharedPtr<Aspose::Words::Fields::FieldRef> field = InsertFieldRef(builder, u"MyBookmark", u"", u"\n");
    field->set_IncludeNoteOrComment(true);
    field->set_InsertHyperlink(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" REF  MyBookmark \\f \\h");
    
    // Insert a REF field, and display whether the referenced bookmark is above or below it.
    field = InsertFieldRef(builder, u"MyBookmark", u"The referenced paragraph is ", u" this field.\n");
    field->set_InsertRelativePosition(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" REF  MyBookmark \\p");
    
    // Display the list number of the bookmark as it appears in the document.
    field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's paragraph number is ", u"\n");
    field->set_InsertParagraphNumber(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" REF  MyBookmark \\n");
    
    // Display the bookmark's list number, but with non-delimiter characters, such as the angle brackets, omitted.
    field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's paragraph number, non-delimiters suppressed, is ", u"\n");
    field->set_InsertParagraphNumber(true);
    field->set_SuppressNonDelimiters(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" REF  MyBookmark \\n \\t");
    
    // Move down one list level.
    System::WithLambda::setter_post_increment_wrap(GETTER_SETTER_LVAL_LAMBDA_ARGS(builder->get_ListFormat(), ListLevelNumber));
    builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u">> \x0001");
    
    // Display the list number of the bookmark and the numbers of all the list levels above it.
    field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's full context paragraph number is ", u"\n");
    field->set_InsertParagraphNumberInFullContext(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" REF  MyBookmark \\w");
    
    builder->InsertBreak(BreakType::PageBreak);
    
    // Display the list level numbers between this REF field, and the bookmark that it is referencing.
    field = InsertFieldRef(builder, u"MyBookmark", u"The bookmark's relative paragraph number is ", u"\n");
    field->set_InsertParagraphNumberInRelativeContext(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" REF  MyBookmark \\r");
    
    // At the end of the document, the bookmark will show up as a list item here.
    builder->Writeln(u"List level above bookmark");
    System::WithLambda::setter_post_increment_wrap(GETTER_SETTER_LVAL_LAMBDA_ARGS(builder->get_ListFormat(), ListLevelNumber));
    builder->get_ListFormat()->get_ListLevel()->set_NumberFormat(u">>> \x0002");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.REF.docx");
    TestFieldRef(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.REF.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldRef)
{
    s_instance->FieldRef();
}

} // namespace gtest_test

void ExField::FieldRD()
{
    //ExStart
    //ExFor:FieldRD
    //ExFor:FieldRD.FileName
    //ExFor:FieldRD.IsPathRelative
    //ExSummary:Shows to use the RD field to create a table of contents entries from headings in other documents.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Use a document builder to insert a table of contents,
    // and then add one entry for the table of contents on the following page.
    builder->InsertField(FieldType::FieldTOC, true);
    builder->InsertBreak(BreakType::PageBreak);
    builder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleName(u"Heading 1");
    builder->Writeln(u"TOC entry from within this document");
    
    // Insert an RD field, which references another local file system document in its FileName property.
    // The TOC will also now accept all headings from the referenced document as entries for its table.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldRD>(builder->InsertField(FieldType::FieldRefDoc, true));
    field->set_FileName(get_ArtifactsDir() + u"ReferencedDocument.docx");
    
    ASSERT_EQ(field->GetFieldCode(), System::String::Format(u" RD  {0}ReferencedDocument.docx", get_ArtifactsDir().Replace(u"\\", u"\\\\")));
    
    // Create the document that the RD field is referencing and insert a heading.
    // This heading will show up as an entry in the TOC field in our first document.
    auto referencedDoc = System::MakeObject<Aspose::Words::Document>();
    auto refDocBuilder = System::MakeObject<Aspose::Words::DocumentBuilder>(referencedDoc);
    refDocBuilder->get_CurrentParagraph()->get_ParagraphFormat()->set_StyleName(u"Heading 1");
    refDocBuilder->Writeln(u"TOC entry from referenced document");
    referencedDoc->Save(get_ArtifactsDir() + u"ReferencedDocument.docx");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.RD.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.RD.docx"));
    
    auto fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(fieldToc->get_Result(), System::String(u"TOC entry from within this document\t\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r") + u"TOC entry from referenced document\t1\r");
    
    auto fieldPageRef = System::ExplicitCast<Aspose::Words::Fields::FieldPageRef>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldPageRef, u" PAGEREF _Toc256000000 \\h ", u"2", fieldPageRef);
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldRD>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldRefDoc, System::String::Format(u" RD  {0}ReferencedDocument.docx", get_ArtifactsDir().Replace(u"\\", u"\\\\")), System::String::Empty, field);
    ASSERT_EQ(field->get_FileName(), get_ArtifactsDir().Replace(u"\\", u"\\\\") + u"ReferencedDocument.docx");
    ASSERT_FALSE(field->get_IsPathRelative());
}

namespace gtest_test
{

TEST_F(ExField, DISABLED_FieldRD)
{
    s_instance->FieldRD();
}

} // namespace gtest_test

void ExField::FieldSetRef()
{
    //ExStart
    //ExFor:FieldRef
    //ExFor:FieldRef.BookmarkName
    //ExFor:FieldSet
    //ExFor:FieldSet.BookmarkName
    //ExFor:FieldSet.BookmarkText
    //ExSummary:Shows how to create bookmarked text with a SET field, and then display it in the document using a REF field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Name bookmarked text with a SET field.
    // This field refers to the "bookmark" not a bookmark structure that appears within the text, but a named variable.
    auto fieldSet = System::ExplicitCast<Aspose::Words::Fields::FieldSet>(builder->InsertField(FieldType::FieldSet, false));
    fieldSet->set_BookmarkName(u"MyBookmark");
    fieldSet->set_BookmarkText(u"Hello world!");
    fieldSet->Update();
    
    ASSERT_EQ(fieldSet->GetFieldCode(), u" SET  MyBookmark \"Hello world!\"");
    
    // Refer to the bookmark by name in a REF field and display its contents.
    auto fieldRef = System::ExplicitCast<Aspose::Words::Fields::FieldRef>(builder->InsertField(FieldType::FieldRef, true));
    fieldRef->set_BookmarkName(u"MyBookmark");
    fieldRef->Update();
    
    ASSERT_EQ(fieldRef->GetFieldCode(), u" REF  MyBookmark");
    ASSERT_EQ(fieldRef->get_Result(), u"Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"Field.SET.REF.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SET.REF.docx"));
    
    ASSERT_EQ(doc->get_Range()->get_Bookmarks()->idx_get(0)->get_Text(), u"Hello world!");
    
    fieldSet = System::ExplicitCast<Aspose::Words::Fields::FieldSet>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldSet, u" SET  MyBookmark \"Hello world!\"", u"Hello world!", fieldSet);
    ASSERT_EQ(fieldSet->get_BookmarkName(), u"MyBookmark");
    ASSERT_EQ(fieldSet->get_BookmarkText(), u"Hello world!");
    
    TestUtil::VerifyField(FieldType::FieldRef, u" REF  MyBookmark", u"Hello world!", fieldRef);
    ASSERT_EQ(fieldRef->get_Result(), u"Hello world!");
}

namespace gtest_test
{

TEST_F(ExField, FieldSetRef)
{
    s_instance->FieldSetRef();
}

} // namespace gtest_test

void ExField::FieldTemplate()
{
    //ExStart
    //ExFor:FieldTemplate
    //ExFor:FieldTemplate.IncludeFullPath
    //ExFor:FieldOptions.TemplateName
    //ExSummary:Shows how to use a TEMPLATE field to display the local file system location of a document's template.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // We can set a template name using by the fields. This property is used when the "doc.AttachedTemplate" is empty.
    // If this property is empty the default template file name "Normal.dotm" is used.
    doc->get_FieldOptions()->set_TemplateName(System::String::Empty);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldTemplate>(builder->InsertField(FieldType::FieldTemplate, false));
    ASSERT_EQ(field->GetFieldCode(), u" TEMPLATE ");
    
    builder->Writeln();
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTemplate>(builder->InsertField(FieldType::FieldTemplate, false));
    field->set_IncludeFullPath(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" TEMPLATE  \\p");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.TEMPLATE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.TEMPLATE.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTemplate>(doc->get_Range()->get_Fields()->idx_get(0));
    ASSERT_EQ(field->GetFieldCode(), u" TEMPLATE ");
    ASSERT_EQ(field->get_Result(), u"Normal.dotm");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTemplate>(doc->get_Range()->get_Fields()->idx_get(1));
    ASSERT_EQ(field->GetFieldCode(), u" TEMPLATE  \\p");
    ASSERT_EQ(field->get_Result(), u"Normal.dotm");
}

namespace gtest_test
{

TEST_F(ExField, FieldTemplate)
{
    s_instance->FieldTemplate();
}

} // namespace gtest_test

void ExField::FieldSymbol()
{
    //ExStart
    //ExFor:FieldSymbol
    //ExFor:FieldSymbol.CharacterCode
    //ExFor:FieldSymbol.DontAffectsLineSpacing
    //ExFor:FieldSymbol.FontName
    //ExFor:FieldSymbol.FontSize
    //ExFor:FieldSymbol.IsAnsi
    //ExFor:FieldSymbol.IsShiftJis
    //ExFor:FieldSymbol.IsUnicode
    //ExSummary:Shows how to use the SYMBOL field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Below are three ways to use a SYMBOL field to display a single character.
    // 1 -  Add a SYMBOL field which displays the © (Copyright) symbol, specified by an ANSI character code:
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));
    
    // The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol.
    field->set_CharacterCode(System::Convert::ToString(0x00a9));
    field->set_IsAnsi(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" SYMBOL  169 \\a");
    
    builder->Writeln(u" Line 1");
    
    // 2 -  Add a SYMBOL field which displays the ∞ (Infinity) symbol, and modify its appearance:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));
    
    // In Unicode, the infinity symbol occupies the "221E" code.
    field->set_CharacterCode(System::Convert::ToString(0x221E));
    field->set_IsUnicode(true);
    
    // Change the font of our symbol after using the Windows Character Map
    // to ensure that the font can represent that symbol.
    field->set_FontName(u"Calibri");
    field->set_FontSize(u"24");
    
    // We can set this flag for tall symbols to make them not push down the rest of the text on their line.
    field->set_DontAffectsLineSpacing(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h");
    
    builder->Writeln(u"Line 2");
    
    // 3 -  Add a SYMBOL field which displays the あ character,
    // with a font that supports Shift-JIS (Windows-932) codepage:
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(builder->InsertField(FieldType::FieldSymbol, true));
    field->set_FontName(u"MS Gothic");
    field->set_CharacterCode(System::Convert::ToString(0x82A0));
    field->set_IsShiftJis(true);
    
    ASSERT_EQ(field->GetFieldCode(), u" SYMBOL  33440 \\f \"MS Gothic\" \\j");
    
    builder->Write(u"Line 3");
    
    doc->Save(get_ArtifactsDir() + u"Field.SYMBOL.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SYMBOL.docx"));
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL  169 \\a", System::String::Empty, field);
    ASSERT_EQ(field->get_CharacterCode(), System::Convert::ToString(0x00a9));
    ASSERT_TRUE(field->get_IsAnsi());
    ASSERT_EQ(field->get_DisplayResult(), u"©");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", System::String::Empty, field);
    ASSERT_EQ(field->get_CharacterCode(), System::Convert::ToString(0x221E));
    ASSERT_EQ(field->get_FontName(), u"Calibri");
    ASSERT_EQ(field->get_FontSize(), u"24");
    ASSERT_TRUE(field->get_IsUnicode());
    ASSERT_TRUE(field->get_DontAffectsLineSpacing());
    ASSERT_EQ(field->get_DisplayResult(), u"∞");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldSymbol>(doc->get_Range()->get_Fields()->idx_get(2));
    
    TestUtil::VerifyField(FieldType::FieldSymbol, u" SYMBOL  33440 \\f \"MS Gothic\" \\j", System::String::Empty, field);
    ASSERT_EQ(field->get_CharacterCode(), System::Convert::ToString(0x82A0));
    ASSERT_EQ(field->get_FontName(), u"MS Gothic");
    ASSERT_TRUE(field->get_IsShiftJis());
}

namespace gtest_test
{

TEST_F(ExField, FieldSymbol)
{
    s_instance->FieldSymbol();
}

} // namespace gtest_test

void ExField::FieldTitle()
{
    //ExStart
    //ExFor:FieldTitle
    //ExFor:FieldTitle.Text
    //ExSummary:Shows how to use the TITLE field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Set a value for the "Title" built-in document property.
    doc->get_BuiltInDocumentProperties()->set_Title(u"My Title");
    
    // We can use the TITLE field to display the value of this property in the document.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldTitle>(builder->InsertField(FieldType::FieldTitle, false));
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" TITLE ");
    ASSERT_EQ(field->get_Result(), u"My Title");
    
    // Setting a value for the field's Text property,
    // and then updating the field will also overwrite the corresponding built-in property with the new value.
    builder->Writeln();
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTitle>(builder->InsertField(FieldType::FieldTitle, false));
    field->set_Text(u"My New Title");
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" TITLE  \"My New Title\"");
    ASSERT_EQ(field->get_Result(), u"My New Title");
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Title(), u"My New Title");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.TITLE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.TITLE.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Title(), u"My New Title");
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTitle>(doc->get_Range()->get_Fields()->idx_get(0));
    
    TestUtil::VerifyField(FieldType::FieldTitle, u" TITLE ", u"My New Title", field);
    
    field = System::ExplicitCast<Aspose::Words::Fields::FieldTitle>(doc->get_Range()->get_Fields()->idx_get(1));
    
    TestUtil::VerifyField(FieldType::FieldTitle, u" TITLE  \"My New Title\"", u"My New Title", field);
    ASSERT_EQ(field->get_Text(), u"My New Title");
}

namespace gtest_test
{

TEST_F(ExField, FieldTitle)
{
    s_instance->FieldTitle();
}

} // namespace gtest_test

void ExField::FieldTOA()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a TOA field, which will create an entry for each TA field in the document,
    // displaying long citations and page numbers for each entry.
    auto fieldToa = System::ExplicitCast<Aspose::Words::Fields::FieldToa>(builder->InsertField(FieldType::FieldTOA, false));
    
    // Set the entry category for our table. This TOA will now only include TA fields
    // that have a matching value in their EntryCategory property.
    fieldToa->set_EntryCategory(u"1");
    
    // Moreover, the Table of Authorities category at index 1 is "Cases",
    // which will show up as our table's title if we set this variable to true.
    fieldToa->set_UseHeading(true);
    
    // We can further filter TA fields by naming a bookmark that they will need to be within the TOA bounds.
    fieldToa->set_BookmarkName(u"MyBookmark");
    
    // By default, a dotted line page-wide tab appears between the TA field's citation
    // and its page number. We can replace it with any text we put on this property.
    // Inserting a tab character will preserve the original tab.
    fieldToa->set_EntrySeparator(u" \t p.");
    
    // If we have multiple TA entries that share the same long citation,
    // all their respective page numbers will show up on one row.
    // We can use this property to specify a string that will separate their page numbers.
    fieldToa->set_PageNumberListSeparator(u" & p. ");
    
    // We can set this to true to get our table to display the word "passim"
    // if there are five or more page numbers in one row.
    fieldToa->set_UsePassim(true);
    
    // One TA field can refer to a range of pages.
    // We can specify a string here to appear between the start and end page numbers for such ranges.
    fieldToa->set_PageRangeSeparator(u" to ");
    
    // The format from the TA fields will carry over into our table.
    // We can disable this by setting the RemoveEntryFormatting flag.
    fieldToa->set_RemoveEntryFormatting(true);
    builder->get_Font()->set_Color(System::Drawing::Color::get_Green());
    builder->get_Font()->set_Name(u"Arial Black");
    
    ASSERT_EQ(fieldToa->GetFieldCode(), u" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f");
    
    builder->InsertBreak(BreakType::PageBreak);
    
    // This TA field will not appear as an entry in the TOA since it is outside
    // the bookmark's bounds that the TOA's BookmarkName property specifies.
    System::SharedPtr<Aspose::Words::Fields::FieldTA> fieldTA = InsertToaEntry(builder, u"1", u"Source 1");
    
    ASSERT_EQ(fieldTA->GetFieldCode(), u" TA  \\c 1 \\l \"Source 1\"");
    
    // This TA field is inside the bookmark,
    // but the entry category does not match that of the table, so the TA field will not include it.
    builder->StartBookmark(u"MyBookmark");
    fieldTA = InsertToaEntry(builder, u"2", u"Source 2");
    
    // This entry will appear in the table.
    fieldTA = InsertToaEntry(builder, u"1", u"Source 3");
    
    // A TOA table does not display short citations,
    // but we can use them as a shorthand to refer to bulky source names that multiple TA fields reference.
    fieldTA->set_ShortCitation(u"S.3");
    
    ASSERT_EQ(fieldTA->GetFieldCode(), u" TA  \\c 1 \\l \"Source 3\" \\s S.3");
    
    // We can format the page number to make it bold/italic using the following properties.
    // We will still see these effects if we set our table to ignore formatting.
    fieldTA = InsertToaEntry(builder, u"1", u"Source 2");
    fieldTA->set_IsBold(true);
    fieldTA->set_IsItalic(true);
    
    ASSERT_EQ(fieldTA->GetFieldCode(), u" TA  \\c 1 \\l \"Source 2\" \\b \\i");
    
    // We can configure TA fields to get their TOA entries to refer to a range of pages that a bookmark spans across.
    // Note that this entry refers to the same source as the one above to share one row in our table.
    // This row will have the page number of the entry above and the page range of this entry,
    // with the table's page list and page number range separators between page numbers.
    fieldTA = InsertToaEntry(builder, u"1", u"Source 3");
    fieldTA->set_PageRangeBookmarkName(u"MyMultiPageBookmark");
    
    builder->StartBookmark(u"MyMultiPageBookmark");
    builder->InsertBreak(BreakType::PageBreak);
    builder->InsertBreak(BreakType::PageBreak);
    builder->InsertBreak(BreakType::PageBreak);
    builder->EndBookmark(u"MyMultiPageBookmark");
    
    ASSERT_EQ(fieldTA->GetFieldCode(), u" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark");
    
    // If we have enabled the "Passim" feature of our table, having 5 or more TA entries with the same source will invoke it.
    for (int32_t i = 0; i < 5; i++)
    {
        InsertToaEntry(builder, u"1", u"Source 4");
    }
    
    builder->EndBookmark(u"MyBookmark");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.TOA.TA.docx");
    TestFieldTOA(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.TOA.TA.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldTOA)
{
    s_instance->FieldTOA();
}

} // namespace gtest_test

void ExField::FieldAddIn()
{
    //ExStart
    //ExFor:FieldAddIn
    //ExSummary:Shows how to process an ADDIN field.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Field sample - ADDIN.docx"));
    
    // Aspose.Words does not support inserting ADDIN fields, but we can still load and read them.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldAddIn>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->GetFieldCode(), u" ADDIN \"My value\" ");
    //ExEnd
    
    doc = DocumentHelper::SaveOpen(doc);
    
    TestUtil::VerifyField(FieldType::FieldAddin, u" ADDIN \"My value\" ", System::String::Empty, doc->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, FieldAddIn)
{
    s_instance->FieldAddIn();
}

} // namespace gtest_test

void ExField::FieldEditTime()
{
    //ExStart
    //ExFor:FieldEditTime
    //ExSummary:Shows how to use the EDITTIME field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // The EDITTIME field will show, in minutes,
    // the time spent with the document open in a Microsoft Word window.
    builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
    builder->Write(u"You've been editing this document for ");
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldEditTime>(builder->InsertField(FieldType::FieldEditTime, true));
    builder->Writeln(u" minutes.");
    
    // This built in document property tracks the minutes. Microsoft Word uses this property
    // to track the time spent with the document open. We can also edit it ourselves.
    doc->get_BuiltInDocumentProperties()->set_TotalEditingTime(10);
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" EDITTIME ");
    ASSERT_EQ(field->get_Result(), u"10");
    
    // The field does not update itself in real-time, and will also have to be
    // manually updated in Microsoft Word anytime we need an accurate value.
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.EDITTIME.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.EDITTIME.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_TotalEditingTime(), 10);
    
    TestUtil::VerifyField(FieldType::FieldEditTime, u" EDITTIME ", u"10", doc->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, FieldEditTime)
{
    s_instance->FieldEditTime();
}

} // namespace gtest_test

void ExField::FieldEQ()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // An EQ field displays a mathematical equation consisting of one or many elements.
    // Each element takes the following form: [switch][options][arguments].
    // There may be one switch, and several possible options.
    // The arguments are a set of coma-separated values enclosed by round braces.
    
    // Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction".
    // We will pass values 1 and 4 as arguments, and we will not use any options.
    // This field will display a fraction with 1 as the numerator and 4 as the denominator.
    System::SharedPtr<Aspose::Words::Fields::FieldEQ> field = InsertFieldEQ(builder, u"\\f(1,4)");
    
    ASSERT_EQ(field->GetFieldCode(), u" EQ \\f(1,4)");
    
    // One EQ field may contain multiple elements placed sequentially.
    // We can also nest elements inside one another by placing the inner elements
    // inside the argument brackets of outer elements.
    // We can find the full list of switches, along with their uses here:
    // https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/
    
    // Below are applications of nine different EQ field switches that we can use to create different kinds of objects.
    // 1 -  Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing:
    InsertFieldEQ(builder, u"\\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)");
    
    // 2 -  Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces:
    // Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output.
    InsertFieldEQ(builder, u"\\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))");
    
    // 3 -  Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline:
    InsertFieldEQ(builder, u"A \\d \\fo30 \\li() B");
    
    // 4 -  Formula consisting of multiple fractions:
    InsertFieldEQ(builder, u"\\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)");
    
    // 5 -  Integral switch "\i", with a summation symbol:
    InsertFieldEQ(builder, u"\\i \\su(n=1,5,n)");
    
    // 6 -  List switch "\l":
    InsertFieldEQ(builder, u"\\l(1,1,2,3,n,8,13)");
    
    // 7 -  Radical switch "\r", displaying a cubed root of x:
    InsertFieldEQ(builder, u"\\r (3,x)");
    
    // 8 -  Subscript/superscript switch "/s", first as a superscript and then as a subscript:
    InsertFieldEQ(builder, u"\\s \\up8(Superscript) Text \\s \\do8(Subscript)");
    
    // 9 -  Box switch "\x", with lines at the top, bottom, left and right of the input:
    InsertFieldEQ(builder, u"\\x \\to \\bo \\le \\ri(5)");
    
    // Some more complex combinations.
    InsertFieldEQ(builder, u"\\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))");
    InsertFieldEQ(builder, u"\\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)");
    InsertFieldEQ(builder, u"\\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)");
    
    doc->Save(get_ArtifactsDir() + u"Field.EQ.docx");
    TestFieldEQ(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.EQ.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldEQ)
{
    s_instance->FieldEQ();
}

} // namespace gtest_test

void ExField::FieldEQAsOfficeMath()
{
    //ExStart
    //ExFor:FieldEQ
    //ExFor:FieldEQ.AsOfficeMath
    //ExSummary:Shows how to replace the EQ field with Office Math.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Field sample - EQ.docx"));
    System::SharedPtr<Aspose::Words::Fields::FieldEQ> fieldEQ = doc->get_Range()->get_Fields()->LINQ_OfType<System::SharedPtr<Aspose::Words::Fields::FieldEQ>>()->LINQ_First();
    
    System::SharedPtr<Aspose::Words::Math::OfficeMath> officeMath = fieldEQ->AsOfficeMath();
    
    fieldEQ->get_Start()->get_ParentNode()->InsertBefore<System::SharedPtr<Aspose::Words::Math::OfficeMath>>(officeMath, fieldEQ->get_Start());
    fieldEQ->Remove();
    
    doc->Save(get_ArtifactsDir() + u"Field.EQAsOfficeMath.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, FieldEQAsOfficeMath)
{
    s_instance->FieldEQAsOfficeMath();
}

} // namespace gtest_test

void ExField::FieldForms()
{
    //ExStart
    //ExFor:FieldFormCheckBox
    //ExFor:FieldFormDropDown
    //ExFor:FieldFormText
    //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
    // These fields are legacy equivalents of the FormField. We can read, but not create these fields using Aspose.Words.
    // In Microsoft Word, we can insert these fields via the Legacy Tools menu in the Developer tab.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Form fields.docx"));
    
    auto fieldFormCheckBox = System::ExplicitCast<Aspose::Words::Fields::FieldFormCheckBox>(doc->get_Range()->get_Fields()->idx_get(1));
    ASSERT_EQ(fieldFormCheckBox->GetFieldCode(), u" FORMCHECKBOX \u0001");
    
    auto fieldFormDropDown = System::ExplicitCast<Aspose::Words::Fields::FieldFormDropDown>(doc->get_Range()->get_Fields()->idx_get(2));
    ASSERT_EQ(fieldFormDropDown->GetFieldCode(), u" FORMDROPDOWN \u0001");
    
    auto fieldFormText = System::ExplicitCast<Aspose::Words::Fields::FieldFormText>(doc->get_Range()->get_Fields()->idx_get(0));
    ASSERT_EQ(fieldFormText->GetFieldCode(), u" FORMTEXT \u0001");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, FieldForms)
{
    s_instance->FieldForms();
}

} // namespace gtest_test

void ExField::FieldFormula()
{
    //ExStart
    //ExFor:FieldFormula
    //ExSummary:Shows how to use the formula field to display the result of an equation.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // Use a field builder to construct a mathematical equation,
    // then create a formula field to display the equation's result in the document.
    auto fieldBuilder = System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldFormula);
    fieldBuilder->AddArgument(2);
    fieldBuilder->AddArgument(u"*");
    fieldBuilder->AddArgument(5);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldFormula>(fieldBuilder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph()));
    field->Update();
    
    ASSERT_EQ(field->GetFieldCode(), u" = 2 * 5 ");
    ASSERT_EQ(field->get_Result(), u"10");
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.FORMULA.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.FORMULA.docx"));
    
    TestUtil::VerifyField(FieldType::FieldFormula, u" = 2 * 5 ", u"10", doc->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, FieldFormula)
{
    s_instance->FieldFormula();
}

} // namespace gtest_test

void ExField::FieldLastSavedBy()
{
    //ExStart
    //ExFor:FieldLastSavedBy
    //ExSummary:Shows how to use the LASTSAVEDBY field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" built-in property.
    // If we make a document programmatically, this property will be null, and we will need to assign a value.
    doc->get_BuiltInDocumentProperties()->set_LastSavedBy(u"John Doe");
    
    // We can use the LASTSAVEDBY field to display the value of this property in the document.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldLastSavedBy>(builder->InsertField(FieldType::FieldLastSavedBy, true));
    
    ASSERT_EQ(field->GetFieldCode(), u" LASTSAVEDBY ");
    ASSERT_EQ(field->get_Result(), u"John Doe");
    
    doc->Save(get_ArtifactsDir() + u"Field.LASTSAVEDBY.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.LASTSAVEDBY.docx"));
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_LastSavedBy(), u"John Doe");
    TestUtil::VerifyField(FieldType::FieldLastSavedBy, u" LASTSAVEDBY ", u"John Doe", doc->get_Range()->get_Fields()->idx_get(0));
}

namespace gtest_test
{

TEST_F(ExField, FieldLastSavedBy)
{
    s_instance->FieldLastSavedBy();
}

} // namespace gtest_test

void ExField::FieldOcx()
{
    //ExStart
    //ExFor:FieldOcx
    //ExSummary:Shows how to insert an OCX field.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldOcx>(builder->InsertField(FieldType::FieldOcx, true));
    
    ASSERT_EQ(field->GetFieldCode(), u" OCX ");
    //ExEnd
    
    TestUtil::VerifyField(FieldType::FieldOcx, u" OCX ", System::String::Empty, field);
}

namespace gtest_test
{

TEST_F(ExField, FieldOcx)
{
    s_instance->FieldOcx();
}

} // namespace gtest_test

void ExField::FieldPrivate()
{
    // Open a Corel WordPerfect document which we have converted to .docx format.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Field sample - PRIVATE.docx"));
    
    // WordPerfect 5.x/6.x documents like the one we have loaded may contain PRIVATE fields.
    // Microsoft Word preserves PRIVATE fields during load/save operations,
    // but provides no functionality for them.
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldPrivate>(doc->get_Range()->get_Fields()->idx_get(0));
    
    ASSERT_EQ(field->GetFieldCode(), u" PRIVATE \"My value\" ");
    ASSERT_EQ(field->get_Type(), FieldType::FieldPrivate);
    
    // We can also insert PRIVATE fields using a document builder.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->InsertField(FieldType::FieldPrivate, true);
    
    // These fields are not a viable way of protecting sensitive information.
    // Unless backward compatibility with older versions of WordPerfect is essential,
    // we can safely remove these fields. We can do this using a DocumentVisiitor implementation.
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 2);
    
    auto remover = System::MakeObject<Aspose::Words::ApiExamples::ExField::FieldPrivateRemover>();
    doc->Accept(remover);
    
    ASSERT_EQ(remover->GetFieldsRemovedCount(), 2);
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 0);
}

namespace gtest_test
{

TEST_F(ExField, FieldPrivate)
{
    s_instance->FieldPrivate();
}

} // namespace gtest_test

void ExField::FieldSection()
{
    //ExStart
    //ExFor:FieldSection
    //ExFor:FieldSectionPages
    //ExSummary:Shows how to use SECTION and SECTIONPAGES fields to number pages by sections.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
    builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
    
    // A SECTION field displays the number of the section it is in.
    builder->Write(u"Section ");
    auto fieldSection = System::ExplicitCast<Aspose::Words::Fields::FieldSection>(builder->InsertField(FieldType::FieldSection, true));
    
    ASSERT_EQ(fieldSection->GetFieldCode(), u" SECTION ");
    
    // A PAGE field displays the number of the page it is in.
    builder->Write(u"\nPage ");
    auto fieldPage = System::ExplicitCast<Aspose::Words::Fields::FieldPage>(builder->InsertField(FieldType::FieldPage, true));
    
    ASSERT_EQ(fieldPage->GetFieldCode(), u" PAGE ");
    
    // A SECTIONPAGES field displays the number of pages that the section it is in spans across.
    builder->Write(u" of ");
    auto fieldSectionPages = System::ExplicitCast<Aspose::Words::Fields::FieldSectionPages>(builder->InsertField(FieldType::FieldSectionPages, true));
    
    ASSERT_EQ(fieldSectionPages->GetFieldCode(), u" SECTIONPAGES ");
    
    // Move out of the header back into the main document and insert two pages.
    // All these pages will be in the first section. Our fields, which appear once every header,
    // will number the current/total pages of this section.
    builder->MoveToDocumentEnd();
    builder->InsertBreak(BreakType::PageBreak);
    builder->InsertBreak(BreakType::PageBreak);
    
    // We can insert a new section with the document builder like this.
    // This will affect the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers.
    builder->InsertBreak(BreakType::SectionBreakNewPage);
    
    // The PAGE field will keep counting pages across the whole document.
    // We can manually reset its count at each section to keep track of pages section-by-section.
    builder->get_CurrentSection()->get_PageSetup()->set_RestartPageNumbering(true);
    builder->InsertBreak(BreakType::PageBreak);
    
    doc->UpdateFields();
    doc->Save(get_ArtifactsDir() + u"Field.SECTION.SECTIONPAGES.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.SECTION.SECTIONPAGES.docx"));
    
    TestUtil::VerifyField(FieldType::FieldSection, u" SECTION ", u"2", doc->get_Range()->get_Fields()->idx_get(0));
    TestUtil::VerifyField(FieldType::FieldPage, u" PAGE ", u"2", doc->get_Range()->get_Fields()->idx_get(1));
    TestUtil::VerifyField(FieldType::FieldSectionPages, u" SECTIONPAGES ", u"2", doc->get_Range()->get_Fields()->idx_get(2));
}

namespace gtest_test
{

TEST_F(ExField, FieldSection)
{
    s_instance->FieldSection();
}

} // namespace gtest_test

void ExField::FieldTime()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // By default, time is displayed in the "h:mm am/pm" format.
    System::SharedPtr<Aspose::Words::Fields::FieldTime> field = InsertFieldTime(builder, u"");
    
    ASSERT_EQ(field->GetFieldCode(), u" TIME ");
    
    // We can use the \@ flag to change the format of our displayed time.
    field = InsertFieldTime(builder, u"\\@ HHmm");
    
    ASSERT_EQ(field->GetFieldCode(), u" TIME \\@ HHmm");
    
    // We can adjust the format to get TIME field to also display the date, according to the Gregorian calendar.
    field = InsertFieldTime(builder, u"\\@ \"M/d/yyyy h mm:ss am/pm\"");
    
    ASSERT_EQ(field->GetFieldCode(), u" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"");
    
    doc->Save(get_ArtifactsDir() + u"Field.TIME.docx");
    TestFieldTime(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.TIME.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExField, FieldTime)
{
    s_instance->FieldTime();
}

} // namespace gtest_test

void ExField::BidiOutline()
{
    //ExStart
    //ExFor:FieldBidiOutline
    //ExFor:FieldShape
    //ExFor:FieldShape.Text
    //ExFor:ParagraphFormat.Bidi
    //ExSummary:Shows how to create right-to-left language-compatible lists with BIDIOUTLINE fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // The BIDIOUTLINE field numbers paragraphs like the AUTONUM/LISTNUM fields,
    // but is only visible when a right-to-left editing language is enabled, such as Hebrew or Arabic.
    // The following field will display ".1", the RTL equivalent of list number "1.".
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldBidiOutline>(builder->InsertField(FieldType::FieldBidiOutline, true));
    builder->Writeln(u"שלום");
    
    ASSERT_EQ(field->GetFieldCode(), u" BIDIOUTLINE ");
    
    // Add two more BIDIOUTLINE fields, which will display ".2" and ".3".
    builder->InsertField(FieldType::FieldBidiOutline, true);
    builder->Writeln(u"שלום");
    builder->InsertField(FieldType::FieldBidiOutline, true);
    builder->Writeln(u"שלום");
    
    // Set the horizontal text alignment for every paragraph in the document to RTL.
    for (auto&& para : System::IterateOver<Aspose::Words::Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)))
    {
        para->get_ParagraphFormat()->set_Bidi(true);
    }
    
    // If we enable a right-to-left editing language in Microsoft Word, our fields will display numbers.
    // Otherwise, they will display "###".
    doc->Save(get_ArtifactsDir() + u"Field.BIDIOUTLINE.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"Field.BIDIOUTLINE.docx"));
    
    for (auto&& fieldBidiOutline : System::IterateOver(doc->get_Range()->get_Fields()))
    {
        TestUtil::VerifyField(FieldType::FieldBidiOutline, u" BIDIOUTLINE ", System::String::Empty, fieldBidiOutline);
    }
}

namespace gtest_test
{

TEST_F(ExField, BidiOutline)
{
    s_instance->BidiOutline();
}

} // namespace gtest_test

void ExField::Legacy()
{
    //ExStart
    //ExFor:FieldEmbed
    //ExFor:FieldShape
    //ExFor:FieldShape.Text
    //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled during loading.
    // Open a document that was created in Microsoft Word 2003.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Legacy fields.doc"));
    
    // If we open the Word document and press Alt+F9, we will see a SHAPE and an EMBED field.
    // A SHAPE field is the anchor/canvas for an AutoShape object with the "In line with text" wrapping style enabled.
    // An EMBED field has the same function, but for an embedded object,
    // such as a spreadsheet from an external Excel document.
    // However, these fields will not appear in the document's Fields collection.
    ASSERT_EQ(doc->get_Range()->get_Fields()->get_Count(), 0);
    
    // These fields are supported only by old versions of Microsoft Word.
    // The document loading process will convert these fields into Shape objects,
    // which we can access in the document's node collection.
    System::SharedPtr<Aspose::Words::NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
    ASSERT_EQ(shapes->get_Count(), 3);
    
    // The first Shape node corresponds to the SHAPE field in the input document,
    // which is the inline canvas for the AutoShape.
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(0));
    ASSERT_EQ(shape->get_ShapeType(), ShapeType::Image);
    
    // The second Shape node is the AutoShape itself.
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(1));
    ASSERT_EQ(shape->get_ShapeType(), ShapeType::Can);
    
    // The third Shape is what was the EMBED field that contained the external spreadsheet.
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(shapes->idx_get(2));
    ASSERT_EQ(shape->get_ShapeType(), ShapeType::OleObject);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, Legacy)
{
    s_instance->Legacy();
}

} // namespace gtest_test

void ExField::SetFieldIndexFormat()
{
    //ExStart
    //ExFor:FieldIndexFormat
    //ExFor:FieldOptions.FieldIndexFormat
    //ExSummary:Shows how to formatting FieldIndex fields.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Write(u"A");
    builder->InsertBreak(BreakType::LineBreak);
    builder->InsertField(u"XE \"A\"");
    builder->Write(u"B");
    
    builder->InsertField(u" INDEX \\e \" · \" \\h \"A\" \\c \"2\" \\z \"1033\"", nullptr);
    
    doc->get_FieldOptions()->set_FieldIndexFormat(FieldIndexFormat::Fancy);
    doc->UpdateFields();
    
    doc->Save(get_ArtifactsDir() + u"Field.SetFieldIndexFormat.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, SetFieldIndexFormat)
{
    s_instance->SetFieldIndexFormat();
}

} // namespace gtest_test

void ExField::ConditionEvaluationExtensionPoint(System::String fieldCode, int8_t comparisonResult, System::String comparisonError, System::String expectedResult)
{
    const System::String left = u"\"left expression\"";
    const System::String operator_ = u"<>";
    const System::String right = u"\"right expression\"";
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    // Field codes that we use in this example:
    // 1.   " IF {0} {1} {2} \"true argument\" \"false argument\" ".
    // 2.   " COMPARE {0} {1} {2} ".
    System::SharedPtr<Aspose::Words::Fields::Field> field = builder->InsertField(System::String::Format(fieldCode, left, operator_, right), nullptr);
    
    // If the "comparisonResult" is undefined, we create "ComparisonEvaluationResult" with string, instead of bool.
    System::SharedPtr<Aspose::Words::Fields::ComparisonEvaluationResult> result = comparisonResult != -1 ? System::MakeObject<Aspose::Words::Fields::ComparisonEvaluationResult>(comparisonResult == 1) : comparisonError != nullptr ? System::MakeObject<Aspose::Words::Fields::ComparisonEvaluationResult>(comparisonError) : nullptr;
    
    auto evaluator = System::MakeObject<Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator>(result);
    builder->get_Document()->get_FieldOptions()->set_ComparisonExpressionEvaluator(evaluator);
    
    builder->get_Document()->UpdateFields();
    
    ASSERT_EQ(field->get_Result(), expectedResult);
    evaluator->AssertInvocationsCount(1)->AssertInvocationArguments(0, left, operator_, right);
}

namespace gtest_test
{

using ExField_ConditionEvaluationExtensionPoint_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExField::ConditionEvaluationExtensionPoint)>::type;

struct ExField_ConditionEvaluationExtensionPoint : public ExField, public Aspose::Words::ApiExamples::ExField, public ::testing::WithParamInterface<ExField_ConditionEvaluationExtensionPoint_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return
        {
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", static_cast<int8_t>(1), nullptr, u"true argument"),
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", static_cast<int8_t>(0), nullptr, u"false argument"),
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", static_cast<int8_t>(-1), u"Custom Error", u"Custom Error"),
            std::make_tuple(u" IF {0} {1} {2} \"true argument\" \"false argument\" ", static_cast<int8_t>(-1), nullptr, u"true argument"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", static_cast<int8_t>(1), nullptr, u"1"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", static_cast<int8_t>(0), nullptr, u"0"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", static_cast<int8_t>(-1), u"Custom Error", u"Custom Error"),
            std::make_tuple(u" COMPARE {0} {1} {2} ", static_cast<int8_t>(-1), nullptr, u"1"),
        };
    }
};

TEST_P(ExField_ConditionEvaluationExtensionPoint, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ConditionEvaluationExtensionPoint(std::get<0>(params), std::get<1>(params), std::get<2>(params), std::get<3>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExField_ConditionEvaluationExtensionPoint, ::testing::ValuesIn(ExField_ConditionEvaluationExtensionPoint::TestCases()));

} // namespace gtest_test

void ExField::ComparisonExpressionEvaluatorNestedFields()
{
    auto document = System::MakeObject<Aspose::Words::Document>();
    
    System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIf)->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIf)->AddArgument(123)->AddArgument(u">")->AddArgument(666)->AddArgument(u"left greater than right")->AddArgument(u"left less than right"))->AddArgument(u"<>")->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIf)->AddArgument(u"left expression")->AddArgument(u"=")->AddArgument(u"right expression")->AddArgument(u"expression are equal")->AddArgument(u"expression are not equal"))->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIf)->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldArgumentBuilder>()->AddText(u"#")->AddField(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldPage)))->AddArgument(u"=")->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldArgumentBuilder>()->AddText(u"#")->AddField(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldNumPages)))->AddArgument(u"the last page")->AddArgument(u"not the last page"))->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIf)->AddArgument(u"unexpected")->AddArgument(u"=")->AddArgument(u"unexpected")->AddArgument(u"unexpected")->AddArgument(u"unexpected"))->BuildAndInsert(document->get_FirstSection()->get_Body()->get_FirstParagraph());
    
    auto evaluator = System::MakeObject<Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator>(nullptr);
    document->get_FieldOptions()->set_ComparisonExpressionEvaluator(evaluator);
    
    document->UpdateFields();
    
    evaluator->AssertInvocationsCount(4)->AssertInvocationArguments(0, u"123", u">", u"666")->AssertInvocationArguments(1, u"\"left expression\"", u"=", u"\"right expression\"")->AssertInvocationArguments(2, u"left less than right", u"<>", u"expression are not equal")->AssertInvocationArguments(3, u"\"#1\"", u"=", u"\"#1\"");
}

namespace gtest_test
{

TEST_F(ExField, ComparisonExpressionEvaluatorNestedFields)
{
    s_instance->ComparisonExpressionEvaluatorNestedFields();
}

} // namespace gtest_test

void ExField::ComparisonExpressionEvaluatorHeaderFooterFields()
{
    auto document = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(document);
    
    builder->InsertBreak(BreakType::PageBreak);
    builder->InsertBreak(BreakType::PageBreak);
    builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
    
    System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldIf)->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldPage))->AddArgument(u"=")->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldNumPages))->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldArgumentBuilder>()->AddField(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldPage))->AddText(u" / ")->AddField(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldNumPages)))->AddArgument(System::MakeObject<Aspose::Words::Fields::FieldArgumentBuilder>()->AddField(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldPage))->AddText(u" / ")->AddField(System::MakeObject<Aspose::Words::Fields::FieldBuilder>(FieldType::FieldNumPages)))->BuildAndInsert(builder->get_CurrentParagraph());
    
    auto evaluator = System::MakeObject<Aspose::Words::ApiExamples::ExField::ComparisonExpressionEvaluator>(nullptr);
    document->get_FieldOptions()->set_ComparisonExpressionEvaluator(evaluator);
    
    document->UpdateFields();
    
    evaluator->AssertInvocationsCount(3)->AssertInvocationArguments(0, u"1", u"=", u"3")->AssertInvocationArguments(1, u"2", u"=", u"3")->AssertInvocationArguments(2, u"3", u"=", u"3");
}

namespace gtest_test
{

TEST_F(ExField, ComparisonExpressionEvaluatorHeaderFooterFields)
{
    s_instance->ComparisonExpressionEvaluatorHeaderFooterFields();
}

} // namespace gtest_test

void ExField::FieldUpdatingCallbackTest()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->InsertField(u" DATE \\@ \"dddd, d MMMM yyyy\" ");
    builder->InsertField(u" TIME ");
    builder->InsertField(u" REVNUM ");
    builder->InsertField(u" AUTHOR  \"John Doe\" ");
    builder->InsertField(u" SUBJECT \"My Subject\" ");
    builder->InsertField(u" QUOTE \"Hello world!\" ");
    
    auto callback = System::MakeObject<Aspose::Words::ApiExamples::ExField::FieldUpdatingCallback>();
    doc->get_FieldOptions()->set_FieldUpdatingCallback(callback);
    
    doc->UpdateFields();
    
    ASSERT_TRUE(callback->get_FieldUpdatedCalls()->Contains(u"Updating John Doe"));
}

namespace gtest_test
{

TEST_F(ExField, FieldUpdatingCallbackTest)
{
    s_instance->FieldUpdatingCallbackTest();
}

} // namespace gtest_test

void ExField::BibliographySources()
{
    //ExStart:BibliographySources
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:Document.Bibliography
    //ExFor:Bibliography
    //ExFor:Bibliography.Sources
    //ExFor:Source
    //ExFor:Source.#ctor(string, SourceType)
    //ExFor:Source.Title
    //ExFor:Source.AbbreviatedCaseNumber
    //ExFor:Source.AlbumTitle
    //ExFor:Source.BookTitle
    //ExFor:Source.Broadcaster
    //ExFor:Source.BroadcastTitle
    //ExFor:Source.CaseNumber
    //ExFor:Source.ChapterNumber
    //ExFor:Source.City
    //ExFor:Source.Comments
    //ExFor:Source.ConferenceName
    //ExFor:Source.CountryOrRegion
    //ExFor:Source.Court
    //ExFor:Source.Day
    //ExFor:Source.DayAccessed
    //ExFor:Source.Department
    //ExFor:Source.Distributor
    //ExFor:Source.Doi
    //ExFor:Source.Edition
    //ExFor:Source.Guid
    //ExFor:Source.Institution
    //ExFor:Source.InternetSiteTitle
    //ExFor:Source.Issue
    //ExFor:Source.JournalName
    //ExFor:Source.Lcid
    //ExFor:Source.Medium
    //ExFor:Source.Month
    //ExFor:Source.MonthAccessed
    //ExFor:Source.NumberVolumes
    //ExFor:Source.Pages
    //ExFor:Source.PatentNumber
    //ExFor:Source.PeriodicalTitle
    //ExFor:Source.ProductionCompany
    //ExFor:Source.PublicationTitle
    //ExFor:Source.Publisher
    //ExFor:Source.RecordingNumber
    //ExFor:Source.RefOrder
    //ExFor:Source.Reporter
    //ExFor:Source.ShortTitle
    //ExFor:Source.SourceType
    //ExFor:Source.StandardNumber
    //ExFor:Source.StateOrProvince
    //ExFor:Source.Station
    //ExFor:Source.Tag
    //ExFor:Source.Theater
    //ExFor:Source.ThesisType
    //ExFor:Source.Type
    //ExFor:Source.Url
    //ExFor:Source.Version
    //ExFor:Source.Volume
    //ExFor:Source.Year
    //ExFor:Source.YearAccessed
    //ExFor:Source.Contributors
    //ExFor:SourceType
    //ExFor:Contributor
    //ExFor:ContributorCollection
    //ExFor:ContributorCollection.Author
    //ExFor:ContributorCollection.Artist
    //ExFor:ContributorCollection.BookAuthor
    //ExFor:ContributorCollection.Compiler
    //ExFor:ContributorCollection.Composer
    //ExFor:ContributorCollection.Conductor
    //ExFor:ContributorCollection.Counsel
    //ExFor:ContributorCollection.Director
    //ExFor:ContributorCollection.Editor
    //ExFor:ContributorCollection.Interviewee
    //ExFor:ContributorCollection.Interviewer
    //ExFor:ContributorCollection.Inventor
    //ExFor:ContributorCollection.Performer
    //ExFor:ContributorCollection.Producer
    //ExFor:ContributorCollection.Translator
    //ExFor:ContributorCollection.Writer
    //ExFor:PersonCollection
    //ExFor:PersonCollection.Count
    //ExFor:PersonCollection.Item(Int32)
    //ExFor:Person.#ctor(string, string, string)
    //ExFor:Person
    //ExFor:Person.First
    //ExFor:Person.Middle
    //ExFor:Person.Last
    //ExSummary:Shows how to get bibliography sources available in the document.
    auto document = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Bibliography sources.docx"));
    
    System::SharedPtr<Aspose::Words::Bibliography::Bibliography> bibliography = document->get_Bibliography();
    ASSERT_EQ(bibliography->get_Sources()->get_Count(), 12);
    
    // Get default data from bibliography sources.
    System::SharedPtr<Aspose::Words::Bibliography::Source> source = bibliography->get_Sources()->LINQ_FirstOrDefault();
    ASSERT_EQ(source->get_Title(), u"Book 0 (No LCID)");
    ASSERT_EQ(source->get_SourceType(), SourceType::Book);
    ASSERT_EQ(source->get_Contributors()->LINQ_Count(), 3);
    ASSERT_TRUE(System::TestTools::IsNull(source->get_AbbreviatedCaseNumber()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_AlbumTitle()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_BookTitle()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Broadcaster()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_BroadcastTitle()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_CaseNumber()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_ChapterNumber()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Comments()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_ConferenceName()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_CountryOrRegion()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Court()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Day()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_DayAccessed()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Department()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Distributor()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Doi()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Edition()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Guid()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Institution()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_InternetSiteTitle()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Issue()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_JournalName()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Lcid()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Medium()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Month()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_MonthAccessed()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_NumberVolumes()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Pages()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_PatentNumber()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_PeriodicalTitle()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_ProductionCompany()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_PublicationTitle()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Publisher()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_RecordingNumber()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_RefOrder()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Reporter()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_ShortTitle()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_StandardNumber()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_StateOrProvince()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Station()));
    ASSERT_EQ(source->get_Tag(), u"BookNoLCID");
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Theater()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_ThesisType()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Type()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Url()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Version()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Volume()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_Year()));
    ASSERT_TRUE(System::TestTools::IsNull(source->get_YearAccessed()));
    
    // Also, you can create a new source.
    auto newSource = System::MakeObject<Aspose::Words::Bibliography::Source>(u"New source", SourceType::Misc);
    
    System::SharedPtr<Aspose::Words::Bibliography::ContributorCollection> contributors = source->get_Contributors();
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Artist()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_BookAuthor()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Compiler()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Composer()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Conductor()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Counsel()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Director()));
    ASSERT_FALSE(System::TestTools::IsNull(contributors->get_Editor()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Interviewee()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Interviewer()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Inventor()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Performer()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Producer()));
    ASSERT_FALSE(System::TestTools::IsNull(contributors->get_Translator()));
    ASSERT_TRUE(System::TestTools::IsNull(contributors->get_Writer()));
    
    System::SharedPtr<Aspose::Words::Bibliography::Contributor> editor = contributors->get_Editor();
    ASSERT_EQ((System::ExplicitCast<Aspose::Words::Bibliography::PersonCollection>(editor))->LINQ_Count(), 2);
    
    auto authors = System::ExplicitCast<Aspose::Words::Bibliography::PersonCollection>(contributors->get_Author());
    ASSERT_EQ(authors->LINQ_Count(), 2);
    
    System::SharedPtr<Aspose::Words::Bibliography::Person> person = authors->idx_get(0);
    ASSERT_EQ(person->get_First(), u"Roxanne");
    ASSERT_EQ(person->get_Middle(), u"Brielle");
    ASSERT_EQ(person->get_Last(), u"Tejeda");
    //ExEnd:BibliographySources
}

namespace gtest_test
{

TEST_F(ExField, BibliographySources)
{
    s_instance->BibliographySources();
}

} // namespace gtest_test

void ExField::BibliographyPersons()
{
    //ExStart
    //ExFor:Person.#ctor(string, string, string)
    //ExFor:PersonCollection.#ctor
    //ExFor:PersonCollection.#ctor(Person[])
    //ExFor:PersonCollection.Add(Person)
    //ExFor:PersonCollection.Contains(Person)
    //ExFor:PersonCollection.Clear
    //ExFor:PersonCollection.Remove(Person)
    //ExFor:PersonCollection.RemoveAt(Int32)
    //ExSummary:Shows how to work with person collection.
    // Create a new person collection.
    auto persons = System::MakeObject<Aspose::Words::Bibliography::PersonCollection>();
    auto person = System::MakeObject<Aspose::Words::Bibliography::Person>(u"Roxanne", u"Brielle", u"Tejeda_updated");
    // Add new person to the collection.
    persons->Add(person);
    ASSERT_EQ(persons->get_Count(), 1);
    // Remove person from the collection if it exists.
    if (persons->Contains(person))
    {
        persons->Remove(person);
    }
    ASSERT_EQ(persons->get_Count(), 0);
    
    // Create person collection with two persons.
    persons = System::MakeObject<Aspose::Words::Bibliography::PersonCollection>(System::MakeArray<System::SharedPtr<Aspose::Words::Bibliography::Person>>({System::MakeObject<Aspose::Words::Bibliography::Person>(u"Roxanne_1", u"Brielle_1", u"Tejeda_1"), System::MakeObject<Aspose::Words::Bibliography::Person>(u"Roxanne_2", u"Brielle_2", u"Tejeda_2")}));
    ASSERT_EQ(persons->get_Count(), 2);
    // Remove person from the collection by the index.
    persons->RemoveAt(0);
    ASSERT_EQ(persons->get_Count(), 1);
    // Remove all persons from the collection.
    persons->Clear();
    ASSERT_EQ(persons->get_Count(), 0);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, BibliographyPersons)
{
    s_instance->BibliographyPersons();
}

} // namespace gtest_test

void ExField::CaptionlessTableOfFiguresLabel()
{
    //ExStart
    //ExFor:FieldToc.CaptionlessTableOfFiguresLabel
    //ExSummary:Shows how to set the name of the sequence identifier.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    auto fieldToc = System::ExplicitCast<Aspose::Words::Fields::FieldToc>(builder->InsertField(FieldType::FieldTOC, true));
    fieldToc->set_CaptionlessTableOfFiguresLabel(u"Test");
    
    ASSERT_EQ(fieldToc->GetFieldCode(), u" TOC  \\a Test");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExField, CaptionlessTableOfFiguresLabel)
{
    s_instance->CaptionlessTableOfFiguresLabel();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
