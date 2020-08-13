// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocument.h"

#include <testing/test_predicates.h>
#include <system/threading/thread.h>
#include <system/text/regularexpressions/regex.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/encoding.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string_comparison.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/scope_guard.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/io/stream.h>
#include <system/io/path.h>
#include <system/io/memory_stream.h>
#include <system/io/file_stream.h>
#include <system/io/file_not_found_exception.h>
#include <system/io/file_mode.h>
#include <system/io/file_info.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/globalization/culture_info.h>
#include <system/func.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <system/collections/ienumerator.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <security/cryptography/x509_certificates/x509_certificate_2.h>
#include <security/cryptography/x509_certificates/x500_distinguished_name.h>
#include <net/web_client.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/size.h>
#include <drawing/rectangle_f.h>
#include <drawing/image_converter.h>
#include <drawing/image.h>
#include <cstdint>
#include <Aspose.Words.Cpp/RW/Ole/VbaProject.h>
#include <Aspose.Words.Cpp/RW/Ole/VbaModuleType.h>
#include <Aspose.Words.Cpp/RW/Ole/VbaModuleCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/VbaModule.h>
#include <Aspose.Words.Cpp/Rendering/ThumbnailGeneratingOptions.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionReference.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionProperty.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtensionBinding.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/WebExtension.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/TaskPane.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionStoreType.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/WebExtensionBindingType.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Enums/TaskPaneDockState.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/WebExtensionPropertyCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/WebExtensionBindingCollection.h>
#include <Aspose.Words.Cpp/Model/WebExtensions/Collections/TaskPaneCollection.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/RowFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/Row.h>
#include <Aspose.Words.Cpp/Model/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Settings/WriteProtection.h>
#include <Aspose.Words.Cpp/Model/Settings/OdsoRecipientDataCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/OdsoRecipientData.h>
#include <Aspose.Words.Cpp/Model/Settings/OdsoFieldMappingType.h>
#include <Aspose.Words.Cpp/Model/Settings/OdsoFieldMapDataCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/OdsoFieldMapData.h>
#include <Aspose.Words.Cpp/Model/Settings/OdsoDataSourceType.h>
#include <Aspose.Words.Cpp/Model/Settings/Odso.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Settings/MailMergeSettings.h>
#include <Aspose.Words.Cpp/Model/Settings/MailMergeMainDocumentType.h>
#include <Aspose.Words.Cpp/Model/Settings/MailMergeDestination.h>
#include <Aspose.Words.Cpp/Model/Settings/MailMergeDataType.h>
#include <Aspose.Words.Cpp/Model/Settings/MailMergeCheckErrors.h>
#include <Aspose.Words.Cpp/Model/Settings/HyphenationOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Saving/ImageSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/IFontSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/FontSavingArgs.h>
#include <Aspose.Words.Cpp/Model/Saving/DownsampleOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/DocumentSplitCriteria.h>
#include <Aspose.Words.Cpp/Model/Saving/DocSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionType.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroupCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionGroup.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>
#include <Aspose.Words.Cpp/Model/Revisions/Revision.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Numbering/NumberStyle.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeChangingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/INodeChangingCallback.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomPartCollection.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomPart.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceType.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingArgs.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLabel.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteType.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Footnotes/FootnoteNumberingRule.h>
#include <Aspose.Words.Cpp/Model/Footnotes/Footnote.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnoteOptions.h>
#include <Aspose.Words.Cpp/Model/Fonts/TableSubstitutionRule.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSubstitutionSettings.h>
#include <Aspose.Words.Cpp/Model/Fonts/FontSettings.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplacingArgs.h>
#include <Aspose.Words.Cpp/Model/FindReplace/ReplaceAction.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldStart.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldSeparator.h>
#include <Aspose.Words.Cpp/Model/Fields/Nodes/FieldEnd.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldUpdateCultureSource.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldDate.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/WarningType.h>
#include <Aspose.Words.Cpp/Model/Document/WarningSource.h>
#include <Aspose.Words.Cpp/Model/Document/WarningInfo.h>
#include <Aspose.Words.Cpp/Model/Document/TxtLoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SubDocument.h>
#include <Aspose.Words.Cpp/Model/Document/SignOptions.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/RevisionsView.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Document/PlainTextDocument.h>
#include <Aspose.Words.Cpp/Model/Document/LoadOptions.h>
#include <Aspose.Words.Cpp/Model/Document/LoadFormat.h>
#include <Aspose.Words.Cpp/Model/Document/LanguagePreferences.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>
#include <Aspose.Words.Cpp/Model/Document/IncorrectPasswordException.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatUtil.h>
#include <Aspose.Words.Cpp/Model/Document/FileFormatInfo.h>
#include <Aspose.Words.Cpp/Model/Document/EditingLanguage.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/Document/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>
#include <Aspose.Words.Cpp/Model/Document/CleanupOptions.h>
#include <Aspose.Words.Cpp/Model/Document/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Comparer/ComparisonTargetType.h>
#include <Aspose.Words.Cpp/Model/Comparer/CompareOptions.h>
#include <Aspose.Words.Cpp/Licensing/License.h>
#include <Aspose.Words.Cpp/Layout/Public/ShowInBalloons.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionTextEffect.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/RevisionColor.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutOptions.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutCollector.h>

#include "TestUtil.h"
#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Properties;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Replacing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::WebExtensions;
using namespace Aspose::Words::Loading;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2967992935u, ::ApiExamples::ExDocument::UseLegacyOrderReplacingCallback, ThisTypeBaseTypesInfo);

Aspose::Words::Replacing::ReplaceAction ExDocument::UseLegacyOrderReplacingCallback::Replacing(SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e)
{
    Matches->Add(e->get_Match()->get_Value());
    return Aspose::Words::Replacing::ReplaceAction::Replace;
}

ExDocument::UseLegacyOrderReplacingCallback::UseLegacyOrderReplacingCallback()
     : Matches(MakeObject<System::Collections::Generic::List<String>>())
{
}

System::Object::shared_members_type ApiExamples::ExDocument::UseLegacyOrderReplacingCallback::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExDocument::UseLegacyOrderReplacingCallback::Matches", this->Matches);

    return result;
}

RTTI_INFO_IMPL_HASH(974002699u, ::ApiExamples::ExDocument::HandleNodeChangingFontChanger, ThisTypeBaseTypesInfo);

void ExDocument::HandleNodeChangingFontChanger::NodeInserted(SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    // Change the font of inserted text contained in the Run nodes
    if (args->get_Node()->get_NodeType() == Aspose::Words::NodeType::Run)
    {
        SharedPtr<Aspose::Words::Font> font = (System::DynamicCast<Aspose::Words::Run>(args->get_Node()))->get_Font();
        font->set_Size(24);
        font->set_Name(u"Arial");
    }
}

void ExDocument::HandleNodeChangingFontChanger::NodeInserting(SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    // Do Nothing
}

void ExDocument::HandleNodeChangingFontChanger::NodeRemoved(SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    // Do Nothing
}

void ExDocument::HandleNodeChangingFontChanger::NodeRemoving(SharedPtr<Aspose::Words::NodeChangingArgs> args)
{
    // Do Nothing
}

RTTI_INFO_IMPL_HASH(4041035362u, ::ApiExamples::ExDocument::HandleFontSaving, ThisTypeBaseTypesInfo);

void ExDocument::HandleFontSaving::FontSaving(SharedPtr<Aspose::Words::Saving::FontSavingArgs> args)
{
    // Print information about fonts
    System::Console::Write(String::Format(u"Font:\t{0}",args->get_FontFamilyName()));
    if (args->get_Bold())
    {
        System::Console::Write(u", bold");
    }
    if (args->get_Italic())
    {
        System::Console::Write(u", italic");
    }
    System::Console::WriteLine(String::Format(u"\nSource:\t{0}, {1} bytes\n",args->get_OriginalFileName(),args->get_OriginalFileSize()));

    ASSERT_TRUE(args->get_IsExportNeeded());
    ASSERT_TRUE(args->get_IsSubsettingNeeded());

    // We can designate where each font will be saved by either specifying a file name, or creating a new stream
    args->set_FontFileName(args->get_OriginalFileName().Split(MakeArray<char16_t>({System::IO::Path::DirectorySeparatorChar}))->LINQ_Last());

    args->set_FontStream(MakeObject<System::IO::FileStream>(ArtifactsDir + args->get_OriginalFileName().Split(MakeArray<char16_t>({System::IO::Path::DirectorySeparatorChar}))->LINQ_Last(), System::IO::FileMode::Create));
    ASSERT_FALSE(args->get_KeepFontStreamOpen());

    // We can access the source document from here also
    ASSERT_TRUE(args->get_Document()->get_OriginalFileName().EndsWith(u"Rendering.docx"));
}

RTTI_INFO_IMPL_HASH(2823058929u, ::ApiExamples::ExDocument::DocumentLoadingWarningCallback, ThisTypeBaseTypesInfo);

void ExDocument::DocumentLoadingWarningCallback::Warning(SharedPtr<Aspose::Words::WarningInfo> info)
{
    System::Console::WriteLine(String::Format(u"Warning: {0}",info->get_WarningType()));
    System::Console::WriteLine(String::Format(u"\tSource: {0}",info->get_Source()));
    System::Console::WriteLine(String::Format(u"\tDescription: {0}",info->get_Description()));
    mWarnings->Add(info);
}

SharedPtr<System::Collections::Generic::List<SharedPtr<Aspose::Words::WarningInfo>>> ExDocument::DocumentLoadingWarningCallback::GetWarnings()
{
    return mWarnings;
}

ExDocument::DocumentLoadingWarningCallback::DocumentLoadingWarningCallback()
     : mWarnings(MakeObject<System::Collections::Generic::List<SharedPtr<WarningInfo>>>())
{
}

System::Object::shared_members_type ApiExamples::ExDocument::DocumentLoadingWarningCallback::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExDocument::DocumentLoadingWarningCallback::mWarnings", this->mWarnings);

    return result;
}

RTTI_INFO_IMPL_HASH(2199000414u, ::ApiExamples::ExDocument::HtmlLinkedResourceLoadingCallback, ThisTypeBaseTypesInfo);

Aspose::Words::Loading::ResourceLoadingAction ExDocument::HtmlLinkedResourceLoadingCallback::ResourceLoading(SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args)
{
    switch (args->get_ResourceType())
    {
        case Aspose::Words::Loading::ResourceType::CssStyleSheet:
            System::Console::WriteLine(String::Format(u"External CSS Stylesheet found upon loading: {0}",args->get_OriginalUri()));
            return Aspose::Words::Loading::ResourceLoadingAction::Default;

        case Aspose::Words::Loading::ResourceType::Image:
        {
            System::Console::WriteLine(String::Format(u"External Image found upon loading: {0}",args->get_OriginalUri()));
            const String newImageFilename = u"Logo.jpg";
            System::Console::WriteLine(String::Format(u"\tImage will be substituted with: {0}",newImageFilename));
            SharedPtr<System::Drawing::Image> newImage = System::Drawing::Image::FromFile(ImageDir + newImageFilename);
            auto converter = MakeObject<System::Drawing::ImageConverter>();
            ArrayPtr<uint8_t> imageBytes = System::DynamicCast<System::Array<uint8_t>>(converter->ConvertTo(newImage, System::ObjectExt::GetType<System::Array<uint8_t>>()));
            args->SetData(imageBytes);
            return Aspose::Words::Loading::ResourceLoadingAction::UserProvided;
        }

        default:
            break;
    }

    return Aspose::Words::Loading::ResourceLoadingAction::Default;
}

void ExDocument::TestLoadOptionsWarningCallback(SharedPtr<System::Collections::Generic::List<SharedPtr<Aspose::Words::WarningInfo>>> warnings)
{
    ASSERT_EQ(Aspose::Words::WarningType::UnexpectedContent, warnings->idx_get(0)->get_WarningType());
    ASSERT_EQ(Aspose::Words::WarningSource::Docx, warnings->idx_get(0)->get_Source());
    ASSERT_EQ(u"3F01", warnings->idx_get(0)->get_Description());

    ASSERT_EQ(Aspose::Words::WarningType::MinorFormattingLoss, warnings->idx_get(1)->get_WarningType());
    ASSERT_EQ(Aspose::Words::WarningSource::Docx, warnings->idx_get(1)->get_Source());
    ASSERT_EQ(u"Import of element 'shapedefaults' is not supported in Docx format by Aspose.Words.", warnings->idx_get(1)->get_Description());

    ASSERT_EQ(Aspose::Words::WarningType::MinorFormattingLoss, warnings->idx_get(2)->get_WarningType());
    ASSERT_EQ(Aspose::Words::WarningSource::Docx, warnings->idx_get(2)->get_Source());
    ASSERT_EQ(u"Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.Words.", warnings->idx_get(2)->get_Description());
}

bool ExDocument::HasParentOfType(SharedPtr<Aspose::Words::Revision> revision, Aspose::Words::NodeType parentType)
{
    SharedPtr<Node> n = revision->get_ParentNode();
    while (n->get_ParentNode() != nullptr)
    {
        if (n->get_NodeType() == parentType)
        {
            return true;
        }
        n = n->get_ParentNode();
    }

    return false;
}

void ExDocument::CheckUseLegacyOrderResults(bool isUseLegacyOrder, SharedPtr<ExDocument::UseLegacyOrderReplacingCallback> callback)
{
    SharedPtr<System::Collections::Generic::List<String>> expected = isUseLegacyOrder ? [&]{ String init_1[] = {u"[tag 1]", u"[tag 2]", u"[tag 3]"}; auto list_1 = MakeObject<System::Collections::Generic::List<String>>(); list_1->AddInitializer(3, init_1); return list_1; }() : [&]{ String init_2[] = {u"[tag 1]", u"[tag 3]", u"[tag 2]"}; auto list_2 = MakeObject<System::Collections::Generic::List<String>>(); list_2->AddInitializer(3, init_2); return list_2; }();
    ASPOSE_ASSERT_EQ(expected, callback->Matches);
}

void ExDocument::TestOdsoEmail(SharedPtr<Aspose::Words::Document> doc)
{
    SharedPtr<Aspose::Words::Settings::MailMergeSettings> settings = doc->get_MailMergeSettings();

    ASSERT_FALSE(settings->get_MailAsAttachment());
    ASSERT_EQ(u"test subject", settings->get_MailSubject());
    ASSERT_EQ(u"Email_Address", settings->get_AddressFieldName());
    ASSERT_EQ(66, settings->get_ActiveRecord());
    ASSERT_EQ(u"SELECT * FROM `Contacts` ", settings->get_Query());

    SharedPtr<Odso> odso = settings->get_Odso();

    ASSERT_EQ(settings->get_ConnectString(), odso->get_UdlConnectString());
    ASSERT_EQ(u"Personal Folders|", odso->get_DataSource());
    ASSERT_EQ(Aspose::Words::Settings::OdsoDataSourceType::Email, odso->get_DataSourceType());
    ASSERT_EQ(u"Contacts", odso->get_TableName());
}

void ExDocument::TestDocPackageCustomParts(SharedPtr<Aspose::Words::Markup::CustomPartCollection> parts)
{
    ASSERT_EQ(3, parts->get_Count());

    ASSERT_EQ(u"/payload/payload_on_package.test", parts->idx_get(0)->get_Name());
    ASSERT_EQ(u"mytest/somedata", parts->idx_get(0)->get_ContentType());
    ASSERT_EQ(u"http://mytest.payload.internal", parts->idx_get(0)->get_RelationshipType());
    ASPOSE_ASSERT_EQ(false, parts->idx_get(0)->get_IsExternal());
    ASSERT_EQ(18, parts->idx_get(0)->get_Data()->get_Length());

    // This part is external and its content is sourced from outside the document
    ASSERT_EQ(u"http://www.aspose.com/Images/aspose-logo.jpg", parts->idx_get(1)->get_Name());
    ASSERT_EQ(u"", parts->idx_get(1)->get_ContentType());
    ASSERT_EQ(u"http://mytest.payload.external", parts->idx_get(1)->get_RelationshipType());
    ASPOSE_ASSERT_EQ(true, parts->idx_get(1)->get_IsExternal());
    ASSERT_EQ(0, parts->idx_get(1)->get_Data()->get_Length());

    ASSERT_EQ(u"http://www.aspose.com/Images/aspose-logo.jpg", parts->idx_get(2)->get_Name());
    ASSERT_EQ(u"", parts->idx_get(2)->get_ContentType());
    ASSERT_EQ(u"http://mytest.payload.external", parts->idx_get(2)->get_RelationshipType());
    ASPOSE_ASSERT_EQ(true, parts->idx_get(2)->get_IsExternal());
    ASSERT_EQ(0, parts->idx_get(2)->get_Data()->get_Length());
}

void ExDocument::TraverseLayoutForward(SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int depth)
{
    do
    {
        PrintCurrentEntity(layoutEnumerator, depth);

        if (layoutEnumerator->MoveFirstChild())
        {
            TraverseLayoutForward(layoutEnumerator, depth + 1);
            layoutEnumerator->MoveParent();
        }
    } while (layoutEnumerator->MoveNext());
}

void ExDocument::TraverseLayoutBackward(SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int depth)
{
    do
    {
        PrintCurrentEntity(layoutEnumerator, depth);

        if (layoutEnumerator->MoveLastChild())
        {
            TraverseLayoutBackward(layoutEnumerator, depth + 1);
            layoutEnumerator->MoveParent();
        }
    } while (layoutEnumerator->MovePrevious());
}

void ExDocument::TraverseLayoutForwardLogical(SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int depth)
{
    do
    {
        PrintCurrentEntity(layoutEnumerator, depth);

        if (layoutEnumerator->MoveFirstChild())
        {
            TraverseLayoutForwardLogical(layoutEnumerator, depth + 1);
            layoutEnumerator->MoveParent();
        }
    } while (layoutEnumerator->MoveNextLogical());
}

void ExDocument::TraverseLayoutBackwardLogical(SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int depth)
{
    do
    {
        PrintCurrentEntity(layoutEnumerator, depth);

        if (layoutEnumerator->MoveLastChild())
        {
            TraverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
            layoutEnumerator->MoveParent();
        }
    } while (layoutEnumerator->MovePreviousLogical());
}

void ExDocument::PrintCurrentEntity(SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int indent)
{
    String tabs(u'\t', indent);

    System::Console::WriteLine(layoutEnumerator->get_Kind() == String::Empty ? String::Format(u"{0}-> Entity type: {1}",tabs,layoutEnumerator->get_Type()) : String::Format(u"{0}-> Entity type & kind: {1}, {2}",tabs,layoutEnumerator->get_Type(),layoutEnumerator->get_Kind()));

    // Only spans can contain text
    if (layoutEnumerator->get_Type() == Aspose::Words::Layout::LayoutEntityType::Span)
    {
        System::Console::WriteLine(String::Format(u"{0}   Span contents: \"{1}\"",tabs,layoutEnumerator->get_Text()));
    }

    System::Drawing::RectangleF leRect = layoutEnumerator->get_Rectangle();
    System::Console::WriteLine(String::Format(u"{0}   Rectangle dimensions {1}x{2}, X={3} Y={4}",tabs,leRect.get_Width(),leRect.get_Height(),leRect.get_X(),leRect.get_Y()));
    System::Console::WriteLine(String::Format(u"{0}   Page {1}",tabs,layoutEnumerator->get_PageIndex()));
}

namespace gtest_test
{

class ExDocument : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExDocument> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDocument>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDocument> ExDocument::s_instance;

} // namespace gtest_test

void ExDocument::LicenseFromFileNoPath()
{
    // This is where the test license is on my development machine.
    String testLicenseFileName = System::IO::Path::Combine(LicenseDir, u"Aspose.Words.Cpp.lic");

    // Copy a license to the bin folder so the example can execute.
    String dstFileName = System::IO::Path::Combine(AssemblyDir, u"Aspose.Words.Cpp.lic");
    System::IO::File::Copy(testLicenseFileName, dstFileName);

    //ExStart
    //ExFor:License
    //ExFor:License.#ctor
    //ExFor:License.SetLicense(String)
    //ExSummary:Aspose.Words will attempt to find the license file in the embedded resources or in the assembly folders.
    auto license = MakeObject<License>();
    license->SetLicense(u"Aspose.Words.Cpp.lic");
    //ExEnd

    // Cleanup by removing the license
    license->SetLicense(u"");
    System::IO::File::Delete(dstFileName);
}

namespace gtest_test
{

TEST_F(ExDocument, LicenseFromFileNoPath)
{
    s_instance->LicenseFromFileNoPath();
}

} // namespace gtest_test

void ExDocument::LicenseFromStream()
{
    // This is where the test license is on my development machine
    String testLicenseFileName = System::IO::Path::Combine(LicenseDir, u"Aspose.Words.Cpp.lic");

    SharedPtr<System::IO::Stream> myStream = System::IO::File::OpenRead(testLicenseFileName);

    {
        auto __finally_guard_0 = ::System::MakeScopeGuard([&myStream]()
        {
            myStream->Close();
        });

        try
        {
            //ExStart
            //ExFor:License.SetLicense(Stream)
            //ExSummary:Shows how to initialize a license from a stream.
            auto license = MakeObject<License>();
            license->SetLicense(myStream);
            //ExEnd
        }
        catch (...)
        {
            throw;
        }
    }
}

namespace gtest_test
{

TEST_F(ExDocument, LicenseFromStream)
{
    s_instance->LicenseFromStream();
}

} // namespace gtest_test

void ExDocument::LoadOptionsCallback()
{
    // Create a new LoadOptions object and set its ResourceLoadingCallback attribute
    // as an instance of our IResourceLoadingCallback implementation
    auto loadOptions = MakeObject<LoadOptions>();
    loadOptions->set_ResourceLoadingCallback(MakeObject<ExDocument::HtmlLinkedResourceLoadingCallback>());

    // When we open an Html document, external resources such as references to CSS stylesheet files and external images
    // will be handled in a custom manner by the loading callback as the document is loaded
    auto doc = MakeObject<Document>(MyDir + u"Images.html", loadOptions);
    doc->Save(ArtifactsDir + u"Document.LoadOptionsCallback.pdf");
}

namespace gtest_test
{

TEST_F(ExDocument, LoadOptionsCallback)
{
    s_instance->LoadOptionsCallback();
}

} // namespace gtest_test

void ExDocument::DocumentCtor()
{
    //ExStart
    //ExFor:Document.#ctor()
    //ExFor:Document.#ctor(String)
    //ExSummary:Shows how to create a blank document.
    // Create a blank document, which will contain a section, body and paragraph by default
    auto doc = MakeObject<Document>();

    // Create a document object from an existing document in the local file system
    doc = MakeObject<Document>(MyDir + u"Document.docx");

    ASSERT_EQ(u"Hello World!", doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DocumentCtor)
{
    s_instance->DocumentCtor();
}

} // namespace gtest_test

void ExDocument::ConvertToPdf()
{
    //ExStart
    //ExFor:Document.#ctor(String)
    //ExFor:Document.Save(String)
    //ExSummary:Shows how to open a document and convert it to .PDF.
    // Open a document that exists in the local file system
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // Save that document as a PDF to another location
    doc->Save(ArtifactsDir + u"Document.ConvertToPdf.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToPdf)
{
    s_instance->ConvertToPdf();
}

} // namespace gtest_test

void ExDocument::OpenAndSaveToFile()
{
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    doc->Save(ArtifactsDir + u"Document.OpenAndSaveToFile.html");
}

namespace gtest_test
{

TEST_F(ExDocument, OpenAndSaveToFile)
{
    s_instance->OpenAndSaveToFile();
}

} // namespace gtest_test

void ExDocument::OpenFromStream()
{
    //ExStart
    //ExFor:Document.#ctor(Stream)
    //ExSummary:Shows how to open a document from a stream.
    // Open the stream. Read only access is enough for Aspose.Words to load a document.
    {
        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Document.docx");
        // Load the entire document into memory and read its contents
        auto doc = MakeObject<Document>(stream);

        ASSERT_EQ(u"Hello World!", doc->GetText().Trim());
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, OpenFromStream)
{
    s_instance->OpenFromStream();
}

} // namespace gtest_test

void ExDocument::OpenFromStreamWithBaseUri()
{
    SharedPtr<Document> doc;

    //ExStart
    //ExFor:Document.#ctor(Stream,LoadOptions)
    //ExFor:LoadOptions.#ctor()
    //ExFor:LoadOptions.BaseUri
    //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
    // Open the stream
    {
        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Document.html");
        // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found
        // Note the Document constructor detects HTML format automatically
        auto loadOptions = MakeObject<LoadOptions>();
        loadOptions->set_BaseUri(ImageDir);

        doc = MakeObject<Document>(stream, loadOptions);
    }
    //ExEnd

    // Save in the DOC format
    doc->Save(ArtifactsDir + u"Document.OpenFromStreamWithBaseUri.doc");

    // Lets make sure the image was imported successfully into a Shape node
    // Get the first shape node in the document
    auto shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    // Verify some properties of the image
    ASSERT_TRUE(shape->get_IsImage());
    ASSERT_FALSE(shape->get_ImageData()->get_ImageBytes() == nullptr);
    ASSERT_NEAR(32.0, ConvertUtil::PointToPixel(shape->get_Width()), 0.01);
    ASSERT_NEAR(32.0, ConvertUtil::PointToPixel(shape->get_Height()), 0.01);
}

namespace gtest_test
{

TEST_F(ExDocument, OpenFromStreamWithBaseUri)
{
    s_instance->OpenFromStreamWithBaseUri();
}

} // namespace gtest_test

void ExDocument::OpenDocumentFromWeb()
{
    //ExStart
    //ExFor:Document.#ctor(Stream)
    //ExSummary:Shows how to retrieve a document from a URL and saves it to disk in a different format.
    // This is the URL address pointing to where to find the document
    const String url = u"https://omextemplates.content.office.net/support/templates/en-us/tf16402488.dotx";

    // The easiest way to load our document from the internet is make use of the
    // System.Net.WebClient class. Create an instance of it and pass the URL
    // to download from.
    {
        auto webClient = MakeObject<System::Net::WebClient>();
        // Download the bytes from the location referenced by the URL
        ArrayPtr<uint8_t> dataBytes = webClient->DownloadData(url);
        ASSERT_TRUE(dataBytes->get_Length() > 0);
        //ExSkip

        // Wrap the bytes representing the document in memory into a MemoryStream object
        {
            auto byteStream = MakeObject<System::IO::MemoryStream>(dataBytes);
            // Load this memory stream into a new Aspose.Words Document
            // The file format of the passed data is inferred from the content of the bytes itself
            // You can load any document format supported by Aspose.Words in the same way
            auto doc = MakeObject<Document>(byteStream);
            ASSERT_TRUE(doc->GetText().Contains(u"First Name last name"));
            //ExSkip

            // Convert the document to any format supported by Aspose.Words and save
            doc->Save(ArtifactsDir + u"Document.OpenDocumentFromWeb.docx");
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, OpenDocumentFromWeb)
{
    s_instance->OpenDocumentFromWeb();
}

} // namespace gtest_test

void ExDocument::InsertHtmlFromWebPage()
{
    //ExStart
    //ExFor:Document.#ctor(Stream, LoadOptions)
    //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
    //ExFor:LoadFormat
    //ExSummary:Shows how to insert the HTML contents from a web page into a new document.
    // The url of the page to load
    const String url = u"https://www.aspose.com/";

    // Create a WebClient object to easily extract the HTML from the page
    auto client = MakeObject<System::Net::WebClient>();
    String pageSource = client->DownloadString(url);
    client->Dispose();

    // Get the HTML as bytes for loading into a stream
    SharedPtr<System::Text::Encoding> encoding = client->get_Encoding();
    ArrayPtr<uint8_t> pageBytes = encoding->GetBytes(pageSource);

    // Load the HTML into a stream
    {
        auto stream = MakeObject<System::IO::MemoryStream>(pageBytes);
        // The baseUri property should be set to ensure any relative img paths are retrieved correctly
        auto options = MakeObject<LoadOptions>(Aspose::Words::LoadFormat::Html, u"", url);

        // Load the HTML document from stream and pass the LoadOptions object
        auto doc = MakeObject<Document>(stream, options);

        // Save the document to the local file system while converting it to .docx
        doc->Save(ArtifactsDir + u"Document.InsertHtmlFromWebPage.docx");
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, InsertHtmlFromWebPage)
{
    s_instance->InsertHtmlFromWebPage();
}

} // namespace gtest_test

void ExDocument::LoadFormat()
{
    //ExStart
    //ExFor:Document.#ctor(String,LoadOptions)
    //ExFor:LoadOptions.LoadFormat
    //ExFor:LoadFormat
    //ExSummary:Shows how to load a document as HTML without automatic file format detection.
    auto loadOptions = MakeObject<LoadOptions>();
    loadOptions->set_LoadFormat(Aspose::Words::LoadFormat::Html);

    auto doc = MakeObject<Document>(MyDir + u"Document.html", loadOptions);
    //ExEnd

    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocument, LoadFormat)
{
    s_instance->LoadFormat();
}

} // namespace gtest_test

void ExDocument::LoadEncrypted()
{
    //ExStart
    //ExFor:Document.#ctor(Stream,LoadOptions)
    //ExFor:Document.#ctor(String,LoadOptions)
    //ExFor:LoadOptions
    //ExFor:LoadOptions.#ctor(String)
    //ExSummary:Shows how to load a Microsoft Word document encrypted with a password.
    // If we try open an encrypted document without the password, an IncorrectPasswordException will be thrown
    // We can construct a LoadOptions object with the correct encryption password
    auto options = MakeObject<LoadOptions>(u"docPassword");

    // Then, we can use that object as a parameter when opening an encrypted document
    auto doc = MakeObject<Document>(MyDir + u"Encrypted.docx", options);
    ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
    //ExSkip

    {
        SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(MyDir + u"Encrypted.docx");
        doc = MakeObject<Document>(stream, options);
        ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
        //ExSkip
    }
    //ExEnd

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&doc]()
    {
        doc = MakeObject<Document>(MyDir + u"Encrypted.docx");
    };

    ASSERT_THROW(_local_func_0(), IncorrectPasswordException);
}

namespace gtest_test
{

TEST_F(ExDocument, LoadEncrypted)
{
    s_instance->LoadEncrypted();
}

} // namespace gtest_test

void ExDocument::ConvertShapeToOfficeMath(bool isConvertShapeToOfficeMath)
{
    //ExStart
    //ExFor:LoadOptions.ConvertShapeToOfficeMath
    //ExSummary:Shows how to convert shapes with EquationXML to Office Math objects.
    auto loadOptions = MakeObject<LoadOptions>();
    // Use 'true/false' values to convert shapes with EquationXML to Office Math objects or not
    loadOptions->set_ConvertShapeToOfficeMath(isConvertShapeToOfficeMath);

    // Specify load option to convert math shapes to office math objects on loading stage
    auto doc = MakeObject<Document>(MyDir + u"Math shapes.docx", loadOptions);
    //ExEnd

    if (isConvertShapeToOfficeMath)
    {
        ASSERT_EQ(16, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
        ASSERT_EQ(34, doc->GetChildNodes(Aspose::Words::NodeType::OfficeMath, true)->get_Count());
    }
    else
    {
        ASSERT_EQ(24, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
        ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::OfficeMath, true)->get_Count());
    }
}

namespace gtest_test
{

using ConvertShapeToOfficeMath_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::ConvertShapeToOfficeMath)>::type;

struct ConvertShapeToOfficeMathVP : public ExDocument, public ApiExamples::ExDocument, public ::testing::WithParamInterface<ConvertShapeToOfficeMath_Args>
{
    static std::vector<ConvertShapeToOfficeMath_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(ConvertShapeToOfficeMathVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->ConvertShapeToOfficeMath(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocument, ConvertShapeToOfficeMathVP, ::testing::ValuesIn(ConvertShapeToOfficeMathVP::TestCases()));

} // namespace gtest_test

void ExDocument::LoadOptionsEncoding()
{
    //ExStart
    //ExFor:LoadOptions.Encoding
    //ExSummary:Shows how to set the encoding with which to open a document.
    // Get the file format info of a file in our local file system
    SharedPtr<FileFormatInfo> fileFormatInfo = FileFormatUtil::DetectFileFormat(MyDir + u"Encoded in UTF-7.txt");

    // A FileFormatInfo object can detect the encoding of the text content in a file, but in some cases it may be ambiguous
    // We know that the above file is encoded in UTF-7, but the text could be valid in others
    ASPOSE_ASSERT_NE(System::Text::Encoding::get_UTF7(), fileFormatInfo->get_Encoding());

    // If we open the document normally, the wrong encoding will be applied,
    // and the content of the document will not be represented correctly
    auto doc = MakeObject<Document>(MyDir + u"Encoded in UTF-7.txt");
    ASSERT_EQ(u"Hello world+ACE-", doc->ToString(Aspose::Words::SaveFormat::Text).Trim());

    // In these cases we can set the Encoding attribute in a LoadOptions object
    // to override the automatically chosen encoding with the one we know to be correct
    auto loadOptions = MakeObject<LoadOptions>();
    loadOptions->set_Encoding(System::Text::Encoding::get_UTF7());
    doc = MakeObject<Document>(MyDir + u"Encoded in UTF-7.txt", loadOptions);

    // This will give us the correct text
    ASSERT_EQ(u"Hello world!", doc->ToString(Aspose::Words::SaveFormat::Text).Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, LoadOptionsEncoding)
{
    s_instance->LoadOptionsEncoding();
}

} // namespace gtest_test

void ExDocument::LoadOptionsFontSettings()
{
    //ExStart
    //ExFor:LoadOptions.FontSettings
    //ExSummary:Shows how to set font settings and apply them during the loading of a document.
    // Create a FontSettings object that will substitute the "Times New Roman" font with the font "Arvo" from our "MyFonts" folder
    auto fontSettings = MakeObject<FontSettings>();
    fontSettings->SetFontsFolder(FontsDir, false);
    fontSettings->get_SubstitutionSettings()->get_TableSubstitution()->AddSubstitutes(u"Times New Roman", MakeArray<String>({u"Arvo"}));

    // Set that FontSettings object as a member of a newly created LoadOptions object
    auto loadOptions = MakeObject<LoadOptions>();
    loadOptions->set_FontSettings(fontSettings);

    // We can now open a document while also passing the LoadOptions object into the constructor so the font substitution occurs upon loading
    auto doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);

    // The effects of our font settings can be observed after rendering
    doc->Save(ArtifactsDir + u"Document.LoadOptionsFontSettings.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, LoadOptionsFontSettings)
{
    s_instance->LoadOptionsFontSettings();
}

} // namespace gtest_test

void ExDocument::LoadOptionsMswVersion()
{
    //ExStart
    //ExFor:LoadOptions.MswVersion
    //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
    // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
    auto loadOptions = MakeObject<LoadOptions>();
    ASSERT_EQ(Aspose::Words::Settings::MsWordVersion::Word2019, loadOptions->get_MswVersion());

    auto doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);
    ASSERT_NEAR(12.95, doc->get_Styles()->get_DefaultParagraphFormat()->get_LineSpacing(), 0.005f);

    // We can change the loading version like this, to Microsoft Word 2007
    loadOptions->set_MswVersion(Aspose::Words::Settings::MsWordVersion::Word2007);

    // This document is missing the default paragraph format style,
    // so when it is opened with either Microsoft Word or Aspose Words, that default style will be regenerated,
    // and will show up in the Styles collection, with values according to Microsoft Word 2007 specifications
    doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);
    ASSERT_NEAR(13.8, doc->get_Styles()->get_DefaultParagraphFormat()->get_LineSpacing(), 0.005f);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, LoadOptionsMswVersion)
{
    s_instance->LoadOptionsMswVersion();
}

} // namespace gtest_test

void ExDocument::LoadOptionsWarningCallback()
{
    // Create a new LoadOptions object and set its WarningCallback attribute as an instance of our IWarningCallback implementation
    auto loadOptions = MakeObject<LoadOptions>();
    loadOptions->set_WarningCallback(MakeObject<ExDocument::DocumentLoadingWarningCallback>());

    // Warnings that occur during loading of the document will now be printed and stored
    auto doc = MakeObject<Document>(MyDir + u"Document.docx", loadOptions);

    SharedPtr<System::Collections::Generic::List<SharedPtr<WarningInfo>>> warnings = (System::StaticCast<ApiExamples::ExDocument::DocumentLoadingWarningCallback>(loadOptions->get_WarningCallback()))->GetWarnings();
    ASSERT_EQ(3, warnings->get_Count());
    TestLoadOptionsWarningCallback(warnings);
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocument, LoadOptionsWarningCallback)
{
    s_instance->LoadOptionsWarningCallback();
}

} // namespace gtest_test

void ExDocument::ConvertToHtml()
{
    //ExStart
    //ExFor:Document.Save(String,SaveFormat)
    //ExFor:SaveFormat
    //ExSummary:Shows how to convert from DOCX to HTML format.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    doc->Save(ArtifactsDir + u"Document.ConvertToHtml.html", Aspose::Words::SaveFormat::Html);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToHtml)
{
    s_instance->ConvertToHtml();
}

} // namespace gtest_test

void ExDocument::ConvertToMhtml()
{
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    doc->Save(ArtifactsDir + u"Document.ConvertToMhtml.mht");
}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToMhtml)
{
    s_instance->ConvertToMhtml();
}

} // namespace gtest_test

void ExDocument::ConvertToTxt()
{
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    doc->Save(ArtifactsDir + u"Document.ConvertToTxt.txt");

}

namespace gtest_test
{

TEST_F(ExDocument, ConvertToTxt)
{
    s_instance->ConvertToTxt();
}

} // namespace gtest_test

void ExDocument::SaveToStream()
{
    //ExStart
    //ExFor:Document.Save(Stream,SaveFormat)
    //ExSummary:Shows how to save a document to a stream.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    {
        auto dstStream = MakeObject<System::IO::MemoryStream>();
        doc->Save(dstStream, Aspose::Words::SaveFormat::Docx);

        // Rewind the stream position back to zero so it is ready for next reader
        dstStream->set_Position(0);
        ASSERT_EQ(u"Hello World!", MakeObject<Document>(dstStream)->GetText().Trim());
        //ExSkip
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SaveToStream)
{
    s_instance->SaveToStream();
}

} // namespace gtest_test

void ExDocument::Doc2EpubSave()
{
    // Open an existing document from disk
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // Save the document in EPUB format
    doc->Save(ArtifactsDir + u"Document.Doc2EpubSave.epub");
}

namespace gtest_test
{

TEST_F(ExDocument, Doc2EpubSave)
{
    s_instance->Doc2EpubSave();
}

} // namespace gtest_test

void ExDocument::Doc2EpubSaveOptions()
{
    //ExStart
    //ExFor:DocumentSplitCriteria
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.#ctor()
    //ExFor:HtmlSaveOptions.Encoding
    //ExFor:HtmlSaveOptions.DocumentSplitCriteria
    //ExFor:HtmlSaveOptions.ExportDocumentProperties
    //ExFor:HtmlSaveOptions.SaveFormat
    //ExFor:SaveOptions
    //ExFor:SaveOptions.SaveFormat
    //ExSummary:Shows how to convert a document to EPUB with save options specified.
    // Open an existing document from disk
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
    // how the output document is saved
    auto saveOptions = MakeObject<HtmlSaveOptions>();
    // Specify the desired encoding
    saveOptions->set_Encoding(System::Text::Encoding::get_UTF8());

    // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB
    // which allows you to limit the size of each HTML part. This is useful for readers which cannot read
    // HTML files greater than a certain size e.g 300kb
    saveOptions->set_DocumentSplitCriteria(Aspose::Words::Saving::DocumentSplitCriteria::HeadingParagraph);

    // Specify that we want to export document properties
    saveOptions->set_ExportDocumentProperties(true);

    // Specify that we want to save in EPUB format
    saveOptions->set_SaveFormat(Aspose::Words::SaveFormat::Epub);

    // Export the document as an EPUB file
    doc->Save(ArtifactsDir + u"Document.Doc2EpubSaveOptions.epub", saveOptions);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, Doc2EpubSaveOptions)
{
    s_instance->Doc2EpubSaveOptions();
}

} // namespace gtest_test

void ExDocument::DownsampleOptions()
{
    //ExStart
    //ExFor:DownsampleOptions
    //ExFor:DownsampleOptions.DownsampleImages
    //ExFor:DownsampleOptions.Resolution
    //ExFor:DownsampleOptions.ResolutionThreshold
    //ExFor:PdfSaveOptions.DownsampleOptions
    //ExSummary:Shows how to change the resolution of images in output pdf documents.
    // Open a document that contains images
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    doc->Save(ArtifactsDir + u"Document.DownsampleOptions.Default.pdf");

    // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
    auto options = MakeObject<PdfSaveOptions>();

    // This conversion will downsample images by default
    ASSERT_TRUE(options->get_DownsampleOptions()->get_DownsampleImages());
    ASSERT_EQ(220, options->get_DownsampleOptions()->get_Resolution());

    // We can set the output resolution to a different value
    // The first two images in the input document will be affected by this
    options->get_DownsampleOptions()->set_Resolution(36);

    // We can set a minimum threshold for downsampling
    // This value will prevent some images in the input document from being downsampled
    options->get_DownsampleOptions()->set_ResolutionThreshold(128);

    doc->Save(ArtifactsDir + u"Document.DownsampleOptions.LowerThreshold.pdf", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DownsampleOptions)
{
    s_instance->DownsampleOptions();
}

} // namespace gtest_test

void ExDocument::SaveHtmlPrettyFormat(bool isPrettyFormat)
{
    //ExStart
    //ExFor:SaveOptions.PrettyFormat
    //ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read
    // If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation
    auto htmlOptions = MakeObject<HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    htmlOptions->set_PrettyFormat(isPrettyFormat);

    doc->Save(ArtifactsDir + u"Document.SaveHtmlPrettyFormat.html", htmlOptions);
    //ExEnd

    String html = System::IO::File::ReadAllText(ArtifactsDir + u"Document.SaveHtmlPrettyFormat.html");

    // Enabling HtmlSaveOptions.PrettyFormat places tabs and newlines in places where it would improve the readability of html source
    ASSERT_TRUE(isPrettyFormat ? html.StartsWith(u"<html>\r\n\t<head>\r\n\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n\t\t") : html.StartsWith(u"<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />"));
}

namespace gtest_test
{

using SaveHtmlPrettyFormat_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::SaveHtmlPrettyFormat)>::type;

struct SaveHtmlPrettyFormatVP : public ExDocument, public ApiExamples::ExDocument, public ::testing::WithParamInterface<SaveHtmlPrettyFormat_Args>
{
    static std::vector<SaveHtmlPrettyFormat_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(SaveHtmlPrettyFormatVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->SaveHtmlPrettyFormat(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocument, SaveHtmlPrettyFormatVP, ::testing::ValuesIn(SaveHtmlPrettyFormatVP::TestCases()));

} // namespace gtest_test

void ExDocument::SaveHtmlWithOptions()
{
    //ExStart
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
    //ExFor:HtmlSaveOptions.ImagesFolder
    //ExSummary:Shows how to set save options before saving a document to HTML.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // This is the directory we want the exported images to be saved to
    String imagesDir = System::IO::Path::Combine(ArtifactsDir, u"SaveHtmlWithOptions");

    // The folder specified needs to exist and should be empty
    if (System::IO::Directory::Exists(imagesDir))
    {
        System::IO::Directory::Delete(imagesDir, true);
    }

    System::IO::Directory::CreateDirectory_(imagesDir);

    // Set an option to export form fields as plain text, not as HTML input elements
    auto options = MakeObject<HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    options->set_ExportTextInputFormFieldAsText(true);
    options->set_ImagesFolder(imagesDir);

    doc->Save(ArtifactsDir + u"Document.SaveHtmlWithOptions.html", options);
    //ExEnd

    // Verify the images were saved to the correct location
    ASSERT_TRUE(System::IO::File::Exists(ArtifactsDir + u"Document.SaveHtmlWithOptions.html"));

    ASSERT_EQ(9, System::IO::Directory::GetFiles(imagesDir)->get_Length());

    System::IO::Directory::Delete(imagesDir, true);
}

namespace gtest_test
{

TEST_F(ExDocument, SaveHtmlWithOptions)
{
    s_instance->SaveHtmlWithOptions();
}

} // namespace gtest_test

void ExDocument::SaveHtmlExportFonts()
{
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // Set the option to export font resources and create and pass the object which implements the handler methods
    auto options = MakeObject<HtmlSaveOptions>(Aspose::Words::SaveFormat::Html);
    options->set_ExportFontResources(true);
    options->set_FontSavingCallback(MakeObject<ExDocument::HandleFontSaving>());

    // The fonts from the input document will now be exported as .ttf files and saved alongside the output document
    doc->Save(ArtifactsDir + u"Document.SaveHtmlExportFonts.html", options);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(String s)> _local_func_1 = [](String s)
    {
        return s.EndsWith(u".ttf");
    };

    ASSERT_EQ(10, System::Array<String>::FindAll(System::IO::Directory::GetFiles(ArtifactsDir), _local_func_1)->get_Length());
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocument, SaveHtmlExportFonts)
{
    s_instance->SaveHtmlExportFonts();
}

} // namespace gtest_test

void ExDocument::FontChangeViaCallback()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set up and pass the object which implements the handler methods
    doc->set_NodeChangingCallback(MakeObject<ExDocument::HandleNodeChangingFontChanger>());

    // Insert sample HTML content
    builder->InsertHtml(u"<p>Hello World</p>");

    doc->Save(ArtifactsDir + u"Document.FontChangeViaCallback.docx");
    doc = MakeObject<Document>(ArtifactsDir + u"Document.FontChangeViaCallback.docx");
    //ExSkip
    auto run = System::DynamicCast<Aspose::Words::Run>(doc->GetChild(Aspose::Words::NodeType::Run, 0, true));
    //ExSkip
    ASPOSE_ASSERT_EQ(24.0, run->get_Font()->get_Size());
    //ExSkip
    ASSERT_EQ(u"Arial", run->get_Font()->get_Name());
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocument, FontChangeViaCallback)
{
    s_instance->FontChangeViaCallback();
}

} // namespace gtest_test

void ExDocument::AppendDocument()
{
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode)
    //ExSummary:Shows how to append a document to the end of another document.
    // The document that the content will be appended to
    auto dstDoc = MakeObject<Document>();
    dstDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Destination document text. ");

    // The document to append
    auto srcDoc = MakeObject<Document>();
    srcDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Source document text. ");

    // Append the source document to the destination document
    // Pass format mode to retain the original formatting of the source document when importing it
    dstDoc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    ASSERT_EQ(2, dstDoc->get_Sections()->get_Count());
    //ExSkip

    // Save the document
    dstDoc->Save(ArtifactsDir + u"Document.AppendDocument.docx");
    //ExEnd

    String outDocText = MakeObject<Document>(ArtifactsDir + u"Document.AppendDocument.docx")->GetText();

    ASSERT_TRUE(outDocText.StartsWith(dstDoc->GetText()));
    ASSERT_TRUE(outDocText.EndsWith(srcDoc->GetText()));
}

namespace gtest_test
{

TEST_F(ExDocument, AppendDocument)
{
    s_instance->AppendDocument();
}

} // namespace gtest_test

void ExDocument::AppendDocumentFromAutomation()
{
    // The document that the other documents will be appended to
    auto doc = MakeObject<Document>();

    // We should call this method to clear this document of any existing content
    doc->RemoveAllChildren();

    const int recordCount = 5;
    for (int i = 1; i <= recordCount; i++)
    {
        auto srcDoc = MakeObject<Document>();

        // Open the document to join.

        // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
        std::function<void()> _local_func_2 = [&srcDoc]()
        {
            srcDoc = MakeObject<Document>(u"C:\\DetailsList.doc");
        };

        ASSERT_THROW(_local_func_2(), System::IO::FileNotFoundException);

        // Append the source document at the end of the destination document
        doc->AppendDocument(srcDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles);

        // In automation you were required to insert a new section break at this point, however in Aspose.Words we
        // don't need to do anything here as the appended document is imported as separate sections already

        // If this is the second document or above being appended then unlink all headers footers in this section
        // from the headers and footers of the previous section
        if (i > 1)
        {

            // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
            std::function<void()> _local_func_3 = [&doc, &i]()
            {
                doc->get_Sections()->idx_get(i)->get_HeadersFooters()->LinkToPrevious(false);
            };

            ASSERT_THROW(_local_func_3(), System::NullReferenceException);
        }
    }
}

namespace gtest_test
{

TEST_F(ExDocument, AppendDocumentFromAutomation)
{
    s_instance->AppendDocumentFromAutomation();
}

} // namespace gtest_test

void ExDocument::ValidateIndividualDocumentSignatures()
{
    //ExStart
    //ExFor:CertificateHolder.Certificate
    //ExFor:Document.DigitalSignatures
    //ExFor:DigitalSignature
    //ExFor:DigitalSignatureCollection
    //ExFor:DigitalSignature.IsValid
    //ExFor:DigitalSignature.Comments
    //ExFor:DigitalSignature.SignTime
    //ExFor:DigitalSignature.SignatureType
    //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
    // Load the document which contains signature
    auto doc = MakeObject<Document>(MyDir + u"Digitally signed.docx");

    for (auto signature : System::IterateOver(doc->get_DigitalSignatures()))
    {
        System::Console::WriteLine(u"*** Signature Found ***");
        System::Console::WriteLine(String(u"Is valid: ") + signature->get_IsValid());
        // This property is available in MS Word documents only
        System::Console::WriteLine(String(u"Reason for signing: ") + signature->get_Comments());
        System::Console::WriteLine(String(u"Signature type: ") + System::ObjectExt::ToString(signature->get_SignatureType()));
        System::Console::WriteLine(String(u"Time of signing: ") + signature->get_SignTime());
        System::Console::WriteLine(String(u"Subject name: ") + signature->get_CertificateHolder()->get_Certificate()->get_SubjectName());
        System::Console::WriteLine(String(u"Issuer name: ") + signature->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name());
        System::Console::WriteLine();
    }
    //ExEnd

    ASSERT_EQ(1, doc->get_DigitalSignatures()->get_Count());

    SharedPtr<Aspose::Words::DigitalSignature> digitalSig = doc->get_DigitalSignatures()->idx_get(0);

    ASSERT_TRUE(digitalSig->get_IsValid());
    ASSERT_EQ(u"Test Sign", digitalSig->get_Comments());
    ASSERT_EQ(u"XmlDsig", System::ObjectExt::ToString(digitalSig->get_SignatureType()));
    ASSERT_TRUE(digitalSig->get_CertificateHolder()->get_Certificate()->get_Subject().Contains(u"Aspose Pty Ltd"));
    ASSERT_TRUE(digitalSig->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name() != nullptr && digitalSig->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name().Contains(u"VeriSign"));
}

namespace gtest_test
{

TEST_F(ExDocument, DISABLED_ValidateIndividualDocumentSignatures)
{
    s_instance->ValidateIndividualDocumentSignatures();
}

} // namespace gtest_test

void ExDocument::DigitalSignature()
{
    //ExStart
    //ExFor:DigitalSignature.CertificateHolder
    //ExFor:DigitalSignature.IssuerName
    //ExFor:DigitalSignature.SubjectName
    //ExFor:DigitalSignatureCollection
    //ExFor:DigitalSignatureCollection.IsValid
    //ExFor:DigitalSignatureCollection.Count
    //ExFor:DigitalSignatureCollection.Item(Int32)
    //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder)
    //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder)
    //ExFor:DigitalSignatureType
    //ExFor:Document.DigitalSignatures
    //ExSummary:Shows how to sign documents with X.509 certificates.
    // Verify that a document isn't signed
    ASSERT_FALSE(FileFormatUtil::DetectFileFormat(MyDir + u"Document.docx")->get_HasDigitalSignature());

    // Create a CertificateHolder object from a PKCS #12 file, which we will use to sign the document
    SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw", nullptr);

    // There are 2 ways of saving a signed copy of a document to the local file system
    // 1: Designate unsigned input and signed output files by filename and sign with the passed CertificateHolder
    DigitalSignatureUtil::Sign(MyDir + u"Document.docx", ArtifactsDir + u"Document.DigitalSignature.docx", certificateHolder, [&]{ auto tmp_0 = MakeObject<SignOptions>(); tmp_0->set_SignTime(System::DateTime::get_Now()); return tmp_0; }());

    ASSERT_TRUE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"Document.DigitalSignature.docx")->get_HasDigitalSignature());

    // 2: Create a stream for the input file and one for the output and create a file, signed with the CertificateHolder, at the file system location determine
    {
        auto inDoc = MakeObject<System::IO::FileStream>(MyDir + u"Document.docx", System::IO::FileMode::Open);
        {
            auto outDoc = MakeObject<System::IO::FileStream>(ArtifactsDir + u"Document.DigitalSignature.docx", System::IO::FileMode::Create);
            DigitalSignatureUtil::Sign(inDoc, outDoc, certificateHolder);
        }
    }

    ASSERT_TRUE(FileFormatUtil::DetectFileFormat(ArtifactsDir + u"Document.DigitalSignature.docx")->get_HasDigitalSignature());

    // Open the signed document and get its digital signature collection
    auto signedDoc = MakeObject<Document>(ArtifactsDir + u"Document.DigitalSignature.docx");
    SharedPtr<DigitalSignatureCollection> digitalSignatureCollection = signedDoc->get_DigitalSignatures();

    // Verify that all of the document's digital signatures are valid and check their details
    ASSERT_TRUE(digitalSignatureCollection->get_IsValid());
    ASSERT_EQ(1, digitalSignatureCollection->get_Count());
    ASSERT_EQ(Aspose::Words::DigitalSignatureType::XmlDsig, digitalSignatureCollection->idx_get(0)->get_SignatureType());
    ASSERT_EQ(u"CN=Morzal.Me", signedDoc->get_DigitalSignatures()->idx_get(0)->get_IssuerName());
    ASSERT_EQ(u"CN=Morzal.Me", signedDoc->get_DigitalSignatures()->idx_get(0)->get_SubjectName());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DigitalSignature)
{
    s_instance->DigitalSignature();
}

} // namespace gtest_test

void ExDocument::AppendAllDocumentsInFolder()
{
    String path = ArtifactsDir + u"Document.AppendAllDocumentsInFolder.doc";

    // Delete the file that was created by the previous run as I don't want to append it again
    if (System::IO::File::Exists(path))
    {
        System::IO::File::Delete(path);
    }

    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode)
    //ExSummary:Shows how to use the AppendDocument method to combine all the documents in a folder to the end of a template document.
    // Lets start with a simple template and append all the documents in a folder to this document
    auto baseDoc = MakeObject<Document>();

    // Add some content to the template
    auto builder = MakeObject<DocumentBuilder>(baseDoc);
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Heading1);
    builder->Writeln(u"Template Document");
    builder->get_ParagraphFormat()->set_StyleIdentifier(Aspose::Words::StyleIdentifier::Normal);
    builder->Writeln(u"Some content here");

    // Gather the files which will be appended to our template document
    // In this case we add the optional parameter to include the search only for files with the ".doc" extension

    auto isDocFile = [](String file)
    {
        return file.EndsWith(u".doc", System::StringComparison::CurrentCultureIgnoreCase);
    };

    auto files = MakeObject<System::Collections::Generic::List<String>>(System::IO::Directory::GetFiles(MyDir, u"*.doc")->LINQ_Where(isDocFile)->LINQ_ToArray());
    ASSERT_EQ(7, files->get_Count());
    //ExSkip

    // The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically
    files->Sort();
    ASSERT_EQ(5, baseDoc->get_Styles()->get_Count());
    //ExSkip
    ASSERT_EQ(1, baseDoc->get_Sections()->get_Count());
    //ExSkip

    // Iterate through every file in the directory and append each one to the end of the template document
    for (auto fileName : files)
    {
        // We have some encrypted test documents in our directory, Aspose.Words can open encrypted documents
        // but only with the correct password. Let's just skip them here for simplicity
        SharedPtr<FileFormatInfo> info = FileFormatUtil::DetectFileFormat(fileName);
        if (info->get_IsEncrypted())
        {
            continue;
        }

        auto subDoc = MakeObject<Document>(fileName);
        baseDoc->AppendDocument(subDoc, Aspose::Words::ImportFormatMode::UseDestinationStyles);
    }

    // Save the combined document to disk
    baseDoc->Save(path);
    //ExEnd

    ASSERT_EQ(7, baseDoc->get_Styles()->get_Count());
    ASSERT_EQ(8, baseDoc->get_Sections()->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocument, AppendAllDocumentsInFolder)
{
    s_instance->AppendAllDocumentsInFolder();
}

} // namespace gtest_test

void ExDocument::JoinRunsWithSameFormatting()
{
    //ExStart
    //ExFor:Document.JoinRunsWithSameFormatting
    //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
    // Open a document which contains adjacent runs of text with identical formatting
    // This can, for example, occur if we edit one paragraph many times
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // Get the number of runs our document contains
    ASSERT_EQ(317, doc->GetChildNodes(Aspose::Words::NodeType::Run, true)->get_Count());

    // We can merge all nearby runs with the same formatting to reduce that number by calling JoinRunsWithSameFormatting()
    // This method will also notify us of the number of run joins that took place
    ASSERT_EQ(121, doc->JoinRunsWithSameFormatting());

    // Get the number of runs after joining, which, together with the number of joins should add up to the original number of runs
    ASSERT_EQ(196, doc->GetChildNodes(Aspose::Words::NodeType::Run, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, JoinRunsWithSameFormatting)
{
    s_instance->JoinRunsWithSameFormatting();
}

} // namespace gtest_test

void ExDocument::DefaultTabStop()
{
    //ExStart
    //ExFor:Document.DefaultTabStop
    //ExFor:ControlChar.Tab
    //ExFor:ControlChar.TabChar
    //ExSummary:Shows how to change default tab positions for the document and inserts text with some tab characters.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set default tab stop to 72 points (1 inch)
    builder->get_Document()->set_DefaultTabStop(72);

    builder->Writeln(String(u"Hello") + ControlChar::Tab() + u"World!");
    builder->Writeln(String(u"Hello") + ControlChar::TabChar + u"World!");
    //ExEnd

    doc = DocumentHelper::SaveOpen(doc);
    ASPOSE_ASSERT_EQ(72, doc->get_DefaultTabStop());
}

namespace gtest_test
{

TEST_F(ExDocument, DefaultTabStop)
{
    s_instance->DefaultTabStop();
}

} // namespace gtest_test

void ExDocument::CloneDocument()
{
    //ExStart
    //ExFor:Document.Clone
    //ExSummary:Shows how to deep clone a document.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    SharedPtr<Document> clone = doc->Clone();

    ASPOSE_ASSERT_NE(doc, clone);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, CloneDocument)
{
    s_instance->CloneDocument();
}

} // namespace gtest_test

void ExDocument::ChangeFieldUpdateCultureSource()
{
    //ExStart
    //ExFor:Document.FieldOptions
    //ExFor:FieldOptions
    //ExFor:FieldOptions.FieldUpdateCultureSource
    //ExFor:FieldUpdateCultureSource
    //ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert two merge fields with German locale
    builder->get_Font()->set_LocaleId(1031);
    builder->InsertField(u"MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
    builder->Write(u" - ");
    builder->InsertField(u"MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

    // Store the current culture in a variable and explicitly set it to US English
    SharedPtr<System::Globalization::CultureInfo> currentCulture = System::Threading::Thread::get_CurrentThread()->get_CurrentCulture();
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"en-US"));

    // Execute a mail merge for the first MERGEFIELD using the current culture (US English) for date formatting
    doc->get_MailMerge()->Execute(MakeArray<String>({u"Date1"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(System::DateTime(2020, 1, 1))}));

    // Execute a mail merge for the second MERGEFIELD using the field's culture (German) for date formatting
    doc->get_FieldOptions()->set_FieldUpdateCultureSource(Aspose::Words::Fields::FieldUpdateCultureSource::FieldCode);
    doc->get_MailMerge()->Execute(MakeArray<String>({u"Date2"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(System::DateTime(2020, 1, 1))}));

    // The first MERGEFIELD has received a date formatted in English, while the second one is in German
    ASSERT_EQ(u"Wednesday, 1 January 2020 - Mittwoch, 1 Januar 2020", doc->get_Range()->get_Text().Trim());

    // Restore the original culture
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(currentCulture);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ChangeFieldUpdateCultureSource)
{
    s_instance->ChangeFieldUpdateCultureSource();
}

} // namespace gtest_test

void ExDocument::DocumentGetTextToString()
{
    //ExStart
    //ExFor:CompositeNode.GetText
    //ExFor:Node.ToString(SaveFormat)
    //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
    auto doc = MakeObject<Document>();

    // Enter a field into the document
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertField(u"MERGEFIELD Field");

    // GetText will retrieve all field codes and special characters
    ASSERT_EQ(u"\u0013MERGEFIELD Field\u0014«Field»\u0015\u000c", doc->GetText());

    // ToString will give us the plaintext version of the document in the save format we put into the parameter
    ASSERT_EQ(u"«Field»\r\n", doc->ToString(Aspose::Words::SaveFormat::Text));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DocumentGetTextToString)
{
    s_instance->DocumentGetTextToString();
}

} // namespace gtest_test

void ExDocument::DocumentByteArray()
{
    // Load the document
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // Create a new memory stream
    auto streamOut = MakeObject<System::IO::MemoryStream>();
    // Save the document to stream
    doc->Save(streamOut, Aspose::Words::SaveFormat::Docx);

    // Convert the document to byte form
    ArrayPtr<uint8_t> docBytes = streamOut->ToArray();

    // We can load the bytes back into a document object
    auto streamIn = MakeObject<System::IO::MemoryStream>(docBytes);

    // Load the stream into a new document object
    auto loadDoc = MakeObject<Document>(streamIn);
    ASSERT_EQ(doc->GetText(), loadDoc->GetText());
}

namespace gtest_test
{

TEST_F(ExDocument, DocumentByteArray)
{
    s_instance->DocumentByteArray();
}

} // namespace gtest_test

void ExDocument::Protect()
{
    //ExStart
    //ExFor:Document.Protect(ProtectionType,String)
    //ExFor:Document.ProtectionType
    //ExFor:Document.Unprotect()
    //ExFor:Document.Unprotect(String)
    //ExSummary:Shows how to protect a document.
    // Create a new document and protect it with a password
    auto doc = MakeObject<Document>();
    doc->Protect(Aspose::Words::ProtectionType::ReadOnly, u"password");
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, doc->get_ProtectionType());

    // If we open this document with Microsoft Word and wish to edit it,
    // we will first need to stop the protection, which can only be done with the password
    doc->Save(ArtifactsDir + u"Document.Protect.docx");

    // Note that the protection only applies to Microsoft Word users opening out document
    // The document can still be opened and edited programmatically without a password, despite its protection status
    // Encryption offers a more robust option for protecting document content
    auto protectedDoc = MakeObject<Document>(ArtifactsDir + u"Document.Protect.docx");
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, protectedDoc->get_ProtectionType());

    auto builder = MakeObject<DocumentBuilder>(protectedDoc);
    builder->Writeln(u"Text added to a protected document.");
    ASSERT_EQ(u"Text added to a protected document.", protectedDoc->get_Range()->get_Text().Trim());
    //ExSkip

    // Documents can have protection removed either with no password, or with the correct password
    doc->Unprotect();
    ASSERT_EQ(Aspose::Words::ProtectionType::NoProtection, doc->get_ProtectionType());

    doc->Protect(Aspose::Words::ProtectionType::ReadOnly, u"newPassword");
    doc->Unprotect(u"wrongPassword");
    //ExSkip
    ASSERT_EQ(Aspose::Words::ProtectionType::ReadOnly, doc->get_ProtectionType());
    //ExSkip
    doc->Unprotect(u"newPassword");
    ASSERT_EQ(Aspose::Words::ProtectionType::NoProtection, doc->get_ProtectionType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, Protect)
{
    s_instance->Protect();
}

} // namespace gtest_test

void ExDocument::DocumentEnsureMinimum()
{
    //ExStart
    //ExFor:Document.EnsureMinimum
    //ExSummary:Shows how to ensure the Document is valid (has the minimum nodes required to be valid).
    auto doc = MakeObject<Document>();

    // Every blank document that we create will contain
    // the minimal set nodes requited for editing; a Section, Body and Paragraph
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // We can remove every node from the document with RemoveAllChildren()
    doc->RemoveAllChildren();
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // EnsureMinimum() can ensure that the document has at least those three nodes
    doc->EnsureMinimum();
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    //ExEnd

    SharedPtr<NodeCollection> nodes = doc->GetChildNodes(Aspose::Words::NodeType::Any, true);

    ASSERT_EQ(Aspose::Words::NodeType::Section, nodes->idx_get(0)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Body, nodes->idx_get(1)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Paragraph, nodes->idx_get(2)->get_NodeType());

    ASSERT_TRUE(nodes->idx_get(1)->get_ParentNode() == nodes->idx_get(0));
    ASSERT_TRUE(nodes->idx_get(2)->get_ParentNode() == nodes->idx_get(1));
}

namespace gtest_test
{

TEST_F(ExDocument, DocumentEnsureMinimum)
{
    s_instance->DocumentEnsureMinimum();
}

} // namespace gtest_test

void ExDocument::RemoveMacrosFromDocument()
{
    //ExStart
    //ExFor:Document.RemoveMacros
    //ExSummary:Shows how to remove all macros from a document.
    // Open a document that contains a VBA project and macros
    auto doc = MakeObject<Document>(MyDir + u"Macro.docm");

    ASSERT_TRUE(doc->get_HasMacros());
    ASSERT_EQ(u"Project", doc->get_VbaProject()->get_Name());
    //ExSkip

    // We can strip the document of this content by calling this method
    doc->RemoveMacros();

    ASSERT_FALSE(doc->get_HasMacros());
    ASSERT_TRUE(doc->get_VbaProject() == nullptr);
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, RemoveMacrosFromDocument)
{
    s_instance->RemoveMacrosFromDocument();
}

} // namespace gtest_test

void ExDocument::UpdateTableLayout()
{
    //ExStart
    //ExFor:Document.UpdateTableLayout
    //ExSummary:Shows how to update the layout of tables in a document.
    // Create a new document and insert a table
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertCell();
    builder->Write(u"Cell 1");
    builder->InsertCell();
    builder->Write(u"Cell 2");
    builder->InsertCell();
    builder->Write(u"Cell 3");

    // Create a SaveOptions object to prepare this document to be saved to .txt
    auto options = MakeObject<TxtSaveOptions>();
    options->set_PreserveTableLayout(true);

    // Previewing the appearance of the document in .txt form shows that the table will not be represented accurately
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));
    //ExSkip
    ASPOSE_ASSERT_EQ(0.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width());
    //ExSkip
    ASSERT_EQ(u"CCC\r\neee\r\nlll\r\nlll\r\n   \r\n123\r\n\r\n", doc->ToString(options));

    // We can call UpdateTableLayout() to fix some of these issues
    doc->UpdateTableLayout();

    ASSERT_EQ(u"Cell 1             Cell 2             Cell 3\r\n\r\n", doc->ToString(options));
    //ExEnd

    ASSERT_NEAR(155.0, table->get_FirstRow()->get_Cells()->idx_get(0)->get_CellFormat()->get_Width(), 2.f);
}

namespace gtest_test
{

TEST_F(ExDocument, UpdateTableLayout)
{
    s_instance->UpdateTableLayout();
}

} // namespace gtest_test

void ExDocument::GetPageCount()
{
    //ExStart
    //ExFor:Document.PageCount
    //ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert text spanning 3 pages
    builder->Write(u"Page 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Page 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Page 3");

    // Get the page count
    ASSERT_EQ(3, doc->get_PageCount());

    // Getting the PageCount property invoked the document's page layout to calculate the value
    // This operation will not need to be re-done when rendering the document to a save format like .pdf,
    // which can save time with larger documents
    doc->Save(ArtifactsDir + u"Document.GetPageCount.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetPageCount)
{
    s_instance->GetPageCount();
}

} // namespace gtest_test

void ExDocument::GetUpdatedPageProperties()
{
    //ExStart
    //ExFor:Document.UpdateWordCount()
    //ExFor:Document.UpdateWordCount(Boolean)
    //ExFor:BuiltInDocumentProperties.Characters
    //ExFor:BuiltInDocumentProperties.Words
    //ExFor:BuiltInDocumentProperties.Paragraphs
    //ExFor:BuiltInDocumentProperties.Lines
    //ExSummary:Shows how to update all list labels in a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add a paragraph of text to the document
    builder->Writeln(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") + u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder->Write(String(u"Ut enim ad minim veniam, ") + u"quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

    // Document metrics are not tracked in code in real time
    ASSERT_EQ(0, doc->get_BuiltInDocumentProperties()->get_Characters());
    ASSERT_EQ(0, doc->get_BuiltInDocumentProperties()->get_Words());
    ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_Paragraphs());
    ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_Lines());

    // We will need to call this method to update them
    doc->UpdateWordCount();

    // Check the values of the properties
    ASSERT_EQ(196, doc->get_BuiltInDocumentProperties()->get_Characters());
    ASSERT_EQ(36, doc->get_BuiltInDocumentProperties()->get_Words());
    ASSERT_EQ(2, doc->get_BuiltInDocumentProperties()->get_Paragraphs());
    ASSERT_EQ(1, doc->get_BuiltInDocumentProperties()->get_Lines());

    // To also get the line count as it would appear in Microsoft Word,
    // we will need to pass "true" to UpdateWordCount()
    doc->UpdateWordCount(true);
    ASSERT_EQ(4, doc->get_BuiltInDocumentProperties()->get_Lines());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetUpdatedPageProperties)
{
    s_instance->GetUpdatedPageProperties();
}

} // namespace gtest_test

void ExDocument::TableStyleToDirectFormatting()
{
    //ExStart
    //ExFor:CompositeNode.GetChild
    //ExFor:Document.ExpandTableStylesToDirectFormatting
    //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
    auto doc = MakeObject<Document>(MyDir + u"Tables.docx");
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(doc->GetChild(Aspose::Words::NodeType::Table, 0, true));

    // First print the color of the cell shading. This should be empty as the current shading
    // is stored in the table style
    double cellShadingBefore = table->get_FirstRow()->get_RowFormat()->get_Height();
    System::Console::WriteLine(String(u"Cell shading before style expansion: ") + cellShadingBefore);

    // Expand table style formatting to direct formatting
    doc->ExpandTableStylesToDirectFormatting();

    // Now print the cell shading after expanding table styles. A blue background pattern color
    // should have been applied from the table style
    double cellShadingAfter = table->get_FirstRow()->get_RowFormat()->get_Height();
    System::Console::WriteLine(String(u"Cell shading after style expansion: ") + cellShadingAfter);

    doc->Save(ArtifactsDir + u"Document.TableStyleToDirectFormatting.docx");
    //ExEnd

    ASPOSE_ASSERT_EQ(0.0, cellShadingBefore);
    ASPOSE_ASSERT_EQ(0.0, cellShadingAfter);
}

namespace gtest_test
{

TEST_F(ExDocument, TableStyleToDirectFormatting)
{
    s_instance->TableStyleToDirectFormatting();
}

} // namespace gtest_test

void ExDocument::GetOriginalFileInfo()
{
    //ExStart
    //ExFor:Document.OriginalFileName
    //ExFor:Document.OriginalLoadFormat
    //ExSummary:Shows how to retrieve the details of the path, filename and LoadFormat of a document from when the document was first loaded into memory.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // This property will return the full path and file name where the document was loaded from
    ASSERT_EQ(MyDir + u"Document.docx", doc->get_OriginalFileName());

    // This is the original LoadFormat of the document
    ASSERT_EQ(Aspose::Words::LoadFormat::Docx, doc->get_OriginalLoadFormat());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetOriginalFileInfo)
{
    s_instance->GetOriginalFileInfo();
}

} // namespace gtest_test

void ExDocument::FootnoteColumns()
{
    //ExStart
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.Columns
    //ExSummary:Shows how to set the number of columns with which the footnotes area is formatted.
    auto doc = MakeObject<Document>(MyDir + u"Footnotes and endnotes.docx");
    ASSERT_EQ(0, doc->get_FootnoteOptions()->get_Columns());
    //ExSkip

    // Let's change number of columns for footnotes on page. If columns value is 0 than footnotes area
    // is formatted with a number of columns based on the number of columns on the displayed page
    doc->get_FootnoteOptions()->set_Columns(2);
    doc->Save(ArtifactsDir + u"Document.FootnoteColumns.docx");
    //ExEnd

    // Assert that number of columns gets correct
    doc = MakeObject<Document>(ArtifactsDir + u"Document.FootnoteColumns.docx");

    ASSERT_EQ(2, doc->get_FirstSection()->get_PageSetup()->get_FootnoteOptions()->get_Columns());
}

namespace gtest_test
{

TEST_F(ExDocument, FootnoteColumns)
{
    s_instance->FootnoteColumns();
}

} // namespace gtest_test

void ExDocument::Footnotes()
{
    //ExStart
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.NumberStyle
    //ExFor:FootnoteOptions.Position
    //ExFor:FootnoteOptions.RestartRule
    //ExFor:FootnoteOptions.StartNumber
    //ExFor:FootnoteNumberingRule
    //ExFor:FootnotePosition
    //ExSummary:Shows how to insert footnotes and edit their appearance.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert 3 paragraphs with a footnote at the end of each one
    builder->Write(u"Text 1. ");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Footnote, u"Footnote 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Footnote, u"Footnote 2");
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Footnote, u"Footnote 3", u"Custom reference mark");

    // Edit the numbering and positioning of footnotes
    doc->get_FootnoteOptions()->set_Position(Aspose::Words::FootnotePosition::BeneathText);
    doc->get_FootnoteOptions()->set_NumberStyle(Aspose::Words::NumberStyle::UppercaseRoman);
    doc->get_FootnoteOptions()->set_RestartRule(Aspose::Words::FootnoteNumberingRule::Continuous);
    doc->get_FootnoteOptions()->set_StartNumber(1);

    doc->Save(ArtifactsDir + u"Document.Footnotes.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.Footnotes.docx");

    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Footnote, true, String::Empty, u"Footnote 1", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Footnote, true, String::Empty, u"Footnote 2", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Footnote, false, u"Custom reference mark", u"Custom reference mark Footnote 3", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 2, true)));
}

namespace gtest_test
{

TEST_F(ExDocument, Footnotes)
{
    s_instance->Footnotes();
}

} // namespace gtest_test

void ExDocument::Endnotes()
{
    //ExStart
    //ExFor:Document.EndnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.NumberStyle
    //ExFor:EndnoteOptions.Position
    //ExFor:EndnoteOptions.RestartRule
    //ExFor:EndnoteOptions.StartNumber
    //ExFor:EndnotePosition
    //ExSummary:Shows how to insert endnotes and edit their appearance.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert 3 paragraphs with an endnote at the end of each one
    builder->Write(u"Text 1. ");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Endnote, u"Endnote 1");
    builder->Write(u"Text 2. ");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Endnote, u"Endnote 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Write(u"Text 3. ");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Endnote, u"Endnote 3", u"Custom reference mark");

    ASSERT_EQ(1, doc->get_EndnoteOptions()->get_StartNumber());
    //ExSkip
    ASSERT_EQ(Aspose::Words::EndnotePosition::EndOfDocument, doc->get_EndnoteOptions()->get_Position());
    //ExSkip
    ASSERT_EQ(Aspose::Words::NumberStyle::LowercaseRoman, doc->get_EndnoteOptions()->get_NumberStyle());
    //ExSkip
    ASSERT_EQ(Aspose::Words::FootnoteNumberingRule::Default, doc->get_EndnoteOptions()->get_RestartRule());
    //ExSkip

    // Edit the numbering and positioning of endnotes
    doc->get_EndnoteOptions()->set_Position(Aspose::Words::EndnotePosition::EndOfDocument);
    doc->get_EndnoteOptions()->set_NumberStyle(Aspose::Words::NumberStyle::UppercaseRoman);
    doc->get_EndnoteOptions()->set_RestartRule(Aspose::Words::FootnoteNumberingRule::Continuous);
    doc->get_EndnoteOptions()->set_StartNumber(1);

    doc->Save(ArtifactsDir + u"Document.Endnotes.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.Endnotes.docx");

    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Endnote, true, String::Empty, u"Endnote 1", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));
    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Endnote, true, String::Empty, u"Endnote 2", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 1, true)));
    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Endnote, false, u"Custom reference mark", u"Custom reference mark Endnote 3", System::DynamicCast<Aspose::Words::Footnote>(doc->GetChild(Aspose::Words::NodeType::Footnote, 2, true)));
}

namespace gtest_test
{

TEST_F(ExDocument, Endnotes)
{
    s_instance->Endnotes();
}

} // namespace gtest_test

void ExDocument::Compare()
{
    //ExStart
    //ExFor:Document.Compare(Document, String, DateTime)
    //ExFor:RevisionCollection.AcceptAll
    //ExSummary:Shows how to apply the compare method to two documents and then use the results.
    auto doc1 = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc1);
    builder->Writeln(u"This is the original document.");

    auto doc2 = MakeObject<Document>();
    builder = MakeObject<DocumentBuilder>(doc2);
    builder->Writeln(u"This is the edited document.");

    // If either document has a revision, an exception will be thrown
    if (doc1->get_Revisions()->get_Count() == 0 && doc2->get_Revisions()->get_Count() == 0)
    {
        doc1->Compare(doc2, u"authorName", System::DateTime::get_Now());
    }

    // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed
    ASSERT_EQ(2, doc1->get_Revisions()->get_Count());
    //ExSkip
    for (auto r : System::IterateOver(doc1->get_Revisions()))
    {
        System::Console::WriteLine(String::Format(u"Revision type: {0}, on a node of type \"{1}\"",r->get_RevisionType(),r->get_ParentNode()->get_NodeType()));
        System::Console::WriteLine(String::Format(u"\tChanged text: \"{0}\"",r->get_ParentNode()->GetText()));
    }

    // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2
    doc1->get_Revisions()->AcceptAll();

    // doc1, when saved, now resembles doc2
    doc1->Save(ArtifactsDir + u"Document.Compare.docx");
    //ExEnd

    doc1 = MakeObject<Document>(ArtifactsDir + u"Document.Compare.docx");
    ASSERT_EQ(0, doc1->get_Revisions()->get_Count());
    ASSERT_EQ(doc2->GetText().Trim(), doc1->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocument, Compare)
{
    s_instance->Compare();
}

} // namespace gtest_test

void ExDocument::CompareDocumentWithRevisions()
{
    auto doc1 = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc1);
    builder->Writeln(u"Hello world! This text is not a revision.");

    auto docWithRevision = MakeObject<Document>();
    builder = MakeObject<DocumentBuilder>(docWithRevision);

    docWithRevision->StartTrackRevisions(u"John Doe");
    builder->Writeln(u"This is a revision.");

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_5 = [&docWithRevision, &doc1]()
    {
        docWithRevision->Compare(doc1, u"John Doe", System::DateTime::get_Now());
    };

    ASSERT_THROW(_local_func_5(), System::InvalidOperationException);
}

namespace gtest_test
{

TEST_F(ExDocument, CompareDocumentWithRevisions)
{
    s_instance->CompareDocumentWithRevisions();
}

} // namespace gtest_test

void ExDocument::CompareOptions()
{
    //ExStart
    //ExFor:CompareOptions
    //ExFor:CompareOptions.IgnoreFormatting
    //ExFor:CompareOptions.IgnoreCaseChanges
    //ExFor:CompareOptions.IgnoreComments
    //ExFor:CompareOptions.IgnoreTables
    //ExFor:CompareOptions.IgnoreFields
    //ExFor:CompareOptions.IgnoreFootnotes
    //ExFor:CompareOptions.IgnoreTextboxes
    //ExFor:CompareOptions.IgnoreHeadersAndFooters
    //ExFor:CompareOptions.Target
    //ExFor:ComparisonTargetType
    //ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
    //ExSummary:Shows how to specify which document shall be used as a target during comparison.
    // Create our original document
    auto docOriginal = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(docOriginal);

    // Insert paragraph text with an endnote
    builder->Writeln(u"Hello world! This is the first paragraph.");
    builder->InsertFootnote(Aspose::Words::FootnoteType::Endnote, u"Original endnote text.");

    // Insert a table
    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"Original cell 1 text");
    builder->InsertCell();
    builder->Write(u"Original cell 2 text");
    builder->EndTable();

    // Insert a textbox
    SharedPtr<Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 150, 20);
    builder->MoveTo(textBox->get_FirstParagraph());
    builder->Write(u"Original textbox contents");

    // Insert a DATE field
    builder->MoveTo(docOriginal->get_FirstSection()->get_Body()->AppendParagraph(u""));
    builder->InsertField(u" DATE ");

    // Insert a comment
    auto newComment = MakeObject<Comment>(docOriginal, u"John Doe", u"J.D.", System::DateTime::get_Now());
    newComment->SetText(u"Original comment.");
    builder->get_CurrentParagraph()->AppendChild(newComment);

    // Insert a header
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"Original header contents.");

    // Create a clone of our document, which we will edit and later compare to the original
    auto docEdited = System::DynamicCast<Aspose::Words::Document>(System::StaticCast<Node>(docOriginal)->Clone(true));
    SharedPtr<Paragraph> firstParagraph = docEdited->get_FirstSection()->get_Body()->get_FirstParagraph();

    // Change the formatting of the first paragraph, change casing of original characters and add text
    firstParagraph->get_Runs()->idx_get(0)->set_Text(u"hello world! this is the first paragraph, after editing.");
    firstParagraph->get_ParagraphFormat()->set_Style(docEdited->get_Styles()->idx_get(Aspose::Words::StyleIdentifier::Heading1));

    // Edit the footnote
    auto footnote = System::DynamicCast<Aspose::Words::Footnote>(docEdited->GetChild(Aspose::Words::NodeType::Footnote, 0, true));
    footnote->get_FirstParagraph()->get_Runs()->idx_get(1)->set_Text(u"Edited endnote text.");

    // Edit the table
    auto table = System::DynamicCast<Aspose::Words::Tables::Table>(docEdited->GetChild(Aspose::Words::NodeType::Table, 0, true));
    table->get_FirstRow()->get_Cells()->idx_get(1)->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited Cell 2 contents");

    // Edit the textbox
    textBox = System::DynamicCast<Aspose::Words::Drawing::Shape>(docEdited->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    textBox->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited textbox contents");

    // Edit the DATE field
    auto fieldDate = System::DynamicCast<Aspose::Words::Fields::FieldDate>(docEdited->get_Range()->get_Fields()->idx_get(0));
    fieldDate->set_UseLunarCalendar(true);

    // Edit the comment
    auto comment = System::DynamicCast<Aspose::Words::Comment>(docEdited->GetChild(Aspose::Words::NodeType::Comment, 0, true));
    comment->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited comment.");

    // Edit the header
    docEdited->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->get_FirstParagraph()->get_Runs()->idx_get(0)->set_Text(u"Edited header contents.");

    // When we compare documents, the differences of the latter document from the former show up as revisions to the former
    // Each edit that we've made above will have its own revision, after we run the Compare method
    // We can compare with a CompareOptions object, which can suppress changes done to certain types of objects within the original document
    // from registering as revisions after the comparison by setting some of these members to "true"
    auto compareOptions = MakeObject<Aspose::Words::CompareOptions>();
    compareOptions->set_IgnoreFormatting(false);
    compareOptions->set_IgnoreCaseChanges(false);
    compareOptions->set_IgnoreComments(false);
    compareOptions->set_IgnoreTables(false);
    compareOptions->set_IgnoreFields(false);
    compareOptions->set_IgnoreFootnotes(false);
    compareOptions->set_IgnoreTextboxes(false);
    compareOptions->set_IgnoreHeadersAndFooters(false);
    compareOptions->set_Target(Aspose::Words::ComparisonTargetType::New);

    docOriginal->Compare(docEdited, u"John Doe", System::DateTime::get_Now(), compareOptions);
    docOriginal->Save(ArtifactsDir + u"Document.CompareOptions.docx");
    //ExEnd

    docOriginal = MakeObject<Document>(ArtifactsDir + u"Document.CompareOptions.docx");

    TestUtil::VerifyFootnote(Aspose::Words::FootnoteType::Endnote, true, String::Empty, u"OriginalEdited endnote text.", System::DynamicCast<Aspose::Words::Footnote>(docOriginal->GetChild(Aspose::Words::NodeType::Footnote, 0, true)));

    // If we set compareOptions to ignore certain types of changes,
    // then revisions done on those types of nodes will not appear in the output document
    // We can tell what kind of node a revision was done on by looking at the NodeType of the revision's parent nodes

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> rev)> _local_func_6 = [](SharedPtr<Revision> rev)
    {
        return rev->get_RevisionType() == Aspose::Words::RevisionType::FormatChange;
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreFormatting(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_6)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> s)> _local_func_7 = [](SharedPtr<Revision> s)
    {
        return s->get_ParentNode()->GetText().Contains(u"hello");
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreCaseChanges(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_7)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> rev)> _local_func_8 = [this](SharedPtr<Revision> rev)
    {
        return HasParentOfType(rev, Aspose::Words::NodeType::Comment);
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreComments(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_8)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> rev)> _local_func_9 = [this](SharedPtr<Revision> rev)
    {
        return HasParentOfType(rev, Aspose::Words::NodeType::Table);
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreTables(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_9)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> rev)> _local_func_10 = [this](SharedPtr<Revision> rev)
    {
        return HasParentOfType(rev, Aspose::Words::NodeType::FieldStart);
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreFields(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_10)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> rev)> _local_func_11 = [this](SharedPtr<Revision> rev)
    {
        return HasParentOfType(rev, Aspose::Words::NodeType::Footnote);
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreFootnotes(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_11)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> rev)> _local_func_12 = [this](SharedPtr<Revision> rev)
    {
        return HasParentOfType(rev, Aspose::Words::NodeType::Shape);
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreTextboxes(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_12)));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<bool(SharedPtr<Revision> rev)> _local_func_13 = [this](SharedPtr<Revision> rev)
    {
        return HasParentOfType(rev, Aspose::Words::NodeType::HeaderFooter);
    };

    ASPOSE_ASSERT_NE(compareOptions->get_IgnoreHeadersAndFooters(), docOriginal->get_Revisions()->LINQ_Any(static_cast<System::Func<SharedPtr<Revision>, bool>>(_local_func_13)));
}

namespace gtest_test
{

TEST_F(ExDocument, CompareOptions)
{
    s_instance->CompareOptions();
}

} // namespace gtest_test

void ExDocument::RemoveExternalSchemaReferences()
{
    //ExStart
    //ExFor:Document.RemoveExternalSchemaReferences
    //ExSummary:Shows how to remove all external XML schema references from a document.
    auto doc = MakeObject<Document>(MyDir + u"External XML schema.docx");

    doc->RemoveExternalSchemaReferences();
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, RemoveExternalSchemaReferences)
{
    s_instance->RemoveExternalSchemaReferences();
}

} // namespace gtest_test

void ExDocument::RemoveUnusedResources()
{
    //ExStart
    //ExFor:Document.Cleanup(CleanupOptions)
    //ExFor:CleanupOptions
    //ExFor:CleanupOptions.UnusedLists
    //ExFor:CleanupOptions.UnusedStyles
    //ExSummary:Shows how to remove all unused styles and lists from a document.
    auto doc = MakeObject<Document>();
    ASSERT_EQ(4, doc->get_Styles()->get_Count());
    //ExSkip

    // Insert some styles into a blank document
    doc->get_Styles()->Add(Aspose::Words::StyleType::List, u"MyListStyle1");
    doc->get_Styles()->Add(Aspose::Words::StyleType::List, u"MyListStyle2");
    doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"MyParagraphStyle1");
    doc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"MyParagraphStyle2");

    // Combined with the built in styles, the document now has 8 styles in total,
    // but all 4 of the ones we added count as unused
    ASSERT_EQ(8, doc->get_Styles()->get_Count());

    // A character style counts as used when the document contains text in that style
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->get_Font()->set_Style(doc->get_Styles()->idx_get(u"MyParagraphStyle1"));
    builder->Writeln(u"Hello world!");

    // A list style is also "used" when there is a list that uses it
    SharedPtr<Aspose::Words::Lists::List> list = doc->get_Lists()->Add(doc->get_Styles()->idx_get(u"MyListStyle1"));
    builder->get_ListFormat()->set_List(list);
    builder->Writeln(u"Item 1");
    builder->Writeln(u"Item 2");

    // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them
    auto cleanupOptions = MakeObject<CleanupOptions>();
    cleanupOptions->set_UnusedLists(true);
    cleanupOptions->set_UnusedStyles(true);

    // We've added 4 styles and used 2 of them, so the other two will be removed when this method is called
    doc->Cleanup(cleanupOptions);
    ASSERT_EQ(6, doc->get_Styles()->get_Count());
    //ExEnd

    doc->get_FirstSection()->get_Body()->RemoveAllChildren();
    doc->Cleanup(cleanupOptions);

    ASSERT_EQ(4, doc->get_Styles()->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocument, RemoveUnusedResources)
{
    s_instance->RemoveUnusedResources();
}

} // namespace gtest_test

void ExDocument::StartTrackRevisions()
{
    //ExStart
    //ExFor:Document.StartTrackRevisions(String)
    //ExFor:Document.StartTrackRevisions(String, DateTime)
    //ExFor:Document.StopTrackRevisions
    //ExSummary:Shows how tracking revisions affects document editing.
    auto doc = MakeObject<Document>();

    // This text will appear as normal text in the document and no revisions will be counted
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->Add(MakeObject<Run>(doc, u"Hello world!"));
    ASSERT_EQ(0, doc->get_Revisions()->get_Count());

    doc->StartTrackRevisions(u"Author");

    // This text will appear as a revision
    // We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
    // on the revision will be the real time when StartTrackRevisions() executes
    doc->get_FirstSection()->get_Body()->AppendParagraph(u"Hello again!");
    ASSERT_EQ(2, doc->get_Revisions()->get_Count());

    // Stopping the tracking of revisions makes this text appear as normal text
    // Revisions are not counted when the document is changed
    doc->StopTrackRevisions();
    doc->get_FirstSection()->get_Body()->AppendParagraph(u"Hello again!");
    ASSERT_EQ(2, doc->get_Revisions()->get_Count());

    // Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called
    // Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all
    doc->StartTrackRevisions(u"Author", System::DateTime(1970, 1, 1));
    doc->get_FirstSection()->get_Body()->AppendParagraph(u"Hello again!");
    ASSERT_EQ(4, doc->get_Revisions()->get_Count());

    doc->Save(ArtifactsDir + u"Document.StartTrackRevisions.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, StartTrackRevisions)
{
    s_instance->StartTrackRevisions();
}

} // namespace gtest_test

void ExDocument::ShowRevisionBalloons()
{
    //ExStart
    //ExFor:RevisionOptions.ShowInBalloons
    //ExSummary:Shows how render tracking changes in balloons
    auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

    // Set option true, if you need render tracking changes in balloons in pdf document,
    // while comments will stay visible
    doc->get_LayoutOptions()->get_RevisionOptions()->set_ShowInBalloons(Aspose::Words::Layout::ShowInBalloons::None);

    // Check that revisions are in balloons
    doc->Save(ArtifactsDir + u"Document.ShowRevisionBalloons.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ShowRevisionBalloons)
{
    s_instance->ShowRevisionBalloons();
}

} // namespace gtest_test

void ExDocument::AcceptAllRevisions()
{
    //ExStart
    //ExFor:Document.AcceptAllRevisions
    //ExSummary:Shows how to accept all tracking changes in the document.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // Start tracking and make some revisions
    doc->StartTrackRevisions(u"Author");
    doc->get_FirstSection()->get_Body()->AppendParagraph(u"Hello world!");
    ASSERT_EQ(2, doc->get_Revisions()->get_Count());
    //ExSkip

    // Revisions will now show up as normal text in the output document
    doc->AcceptAllRevisions();
    doc->Save(ArtifactsDir + u"Document.AcceptAllRevisions.docx");
    ASSERT_EQ(0, doc->get_Revisions()->get_Count());
    //ExSKip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, AcceptAllRevisions)
{
    s_instance->AcceptAllRevisions();
}

} // namespace gtest_test

void ExDocument::GetRevisedPropertiesOfList()
{
    //ExStart
    //ExFor:RevisionsView
    //ExFor:Document.RevisionsView
    //ExSummary:Shows how to get revised version of list label and list level formatting in a document.
    auto doc = MakeObject<Document>(MyDir + u"Revisions at list levels.docx");
    doc->UpdateListLabels();

    // Switch to the revised version of the document
    doc->set_RevisionsView(Aspose::Words::RevisionsView::Final);

    for (auto revision : System::IterateOver(doc->get_Revisions()))
    {
        if (revision->get_ParentNode()->get_NodeType() == Aspose::Words::NodeType::Paragraph)
        {
            auto paragraph = System::DynamicCast<Aspose::Words::Paragraph>(revision->get_ParentNode());

            if (paragraph->get_IsListItem())
            {
                // Print revised version of LabelString and ListLevel
                System::Console::WriteLine(paragraph->get_ListLabel()->get_LabelString());
                System::Console::WriteLine(paragraph->get_ListFormat()->get_ListLevel());
            }
        }
    }
    //ExEnd

    ASSERT_EQ(u"", (System::DynamicCast<Aspose::Words::Paragraph>(doc->get_Revisions()->idx_get(0)->get_ParentNode()))->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"1.", (System::DynamicCast<Aspose::Words::Paragraph>(doc->get_Revisions()->idx_get(1)->get_ParentNode()))->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"a.", (System::DynamicCast<Aspose::Words::Paragraph>(doc->get_Revisions()->idx_get(3)->get_ParentNode()))->get_ListLabel()->get_LabelString());

    doc->set_RevisionsView(Aspose::Words::RevisionsView::Original);

    ASSERT_EQ(u"1.", (System::DynamicCast<Aspose::Words::Paragraph>(doc->get_Revisions()->idx_get(0)->get_ParentNode()))->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"a.", (System::DynamicCast<Aspose::Words::Paragraph>(doc->get_Revisions()->idx_get(1)->get_ParentNode()))->get_ListLabel()->get_LabelString());
    ASSERT_EQ(u"", (System::DynamicCast<Aspose::Words::Paragraph>(doc->get_Revisions()->idx_get(3)->get_ParentNode()))->get_ListLabel()->get_LabelString());
}

namespace gtest_test
{

TEST_F(ExDocument, GetRevisedPropertiesOfList)
{
    s_instance->GetRevisedPropertiesOfList();
}

} // namespace gtest_test

void ExDocument::UpdateThumbnail()
{
    //ExStart
    //ExFor:Document.UpdateThumbnail()
    //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
    //ExFor:ThumbnailGeneratingOptions
    //ExFor:ThumbnailGeneratingOptions.GenerateFromFirstPage
    //ExFor:ThumbnailGeneratingOptions.ThumbnailSize
    //ExSummary:Shows how to update a document's thumbnail.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

    // If we aren't setting the thumbnail via built in document properties,
    // we can set the first page of the document to be the thumbnail in an output .epub like this
    doc->UpdateThumbnail();
    doc->Save(ArtifactsDir + u"Document.UpdateThumbnail.FirstPage.epub");

    // Another way is to use the first image shape found in the document as the thumbnail
    // Insert an image with a builder that we want to use as a thumbnail
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertImage(ImageDir + u"Logo.jpg");

    auto options = MakeObject<ThumbnailGeneratingOptions>();
    ASPOSE_ASSERT_EQ(System::Drawing::Size(600, 900), options->get_ThumbnailSize());
    //ExSKip
    ASSERT_TRUE(options->get_GenerateFromFirstPage());
    //ExSkip
    options->set_ThumbnailSize(System::Drawing::Size(400, 400));
    options->set_GenerateFromFirstPage(false);

    doc->UpdateThumbnail(options);
    doc->Save(ArtifactsDir + u"Document.UpdateThumbnail.FirstImage.epub");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, UpdateThumbnail)
{
    s_instance->UpdateThumbnail();
}

} // namespace gtest_test

void ExDocument::HyphenationOptions()
{
    //ExStart
    //ExFor:Document.HyphenationOptions
    //ExFor:HyphenationOptions
    //ExFor:HyphenationOptions.AutoHyphenation
    //ExFor:HyphenationOptions.ConsecutiveHyphenLimit
    //ExFor:HyphenationOptions.HyphenationZone
    //ExFor:HyphenationOptions.HyphenateCaps
    //ExFor:ParagraphFormat.SuppressAutoHyphens
    //ExSummary:Shows how to configure document hyphenation options.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Set this to insert a page break before this paragraph
    builder->get_Font()->set_Size(24);
    builder->get_ParagraphFormat()->set_SuppressAutoHyphens(false);

    builder->Writeln(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, ") + u"sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    doc->get_HyphenationOptions()->set_AutoHyphenation(true);
    doc->get_HyphenationOptions()->set_ConsecutiveHyphenLimit(2);
    doc->get_HyphenationOptions()->set_HyphenationZone(720);
    // 0.5 inch
    doc->get_HyphenationOptions()->set_HyphenateCaps(true);

    // Each paragraph has this flag that can be set to suppress hyphenation
    ASSERT_FALSE(builder->get_ParagraphFormat()->get_SuppressAutoHyphens());

    doc->Save(ArtifactsDir + u"Document.HyphenationOptions.docx");
    //ExEnd

    ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_AutoHyphenation());
    ASSERT_EQ(2, doc->get_HyphenationOptions()->get_ConsecutiveHyphenLimit());
    ASSERT_EQ(720, doc->get_HyphenationOptions()->get_HyphenationZone());
    ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_HyphenateCaps());

    ASSERT_TRUE(DocumentHelper::CompareDocs(ArtifactsDir + u"Document.HyphenationOptions.docx", GoldsDir + u"Document.HyphenationOptions Gold.docx"));
}

namespace gtest_test
{

TEST_F(ExDocument, HyphenationOptions)
{
    s_instance->HyphenationOptions();
}

} // namespace gtest_test

void ExDocument::HyphenationOptionsDefaultValues()
{
    auto doc = MakeObject<Document>();
    doc = DocumentHelper::SaveOpen(doc);

    ASPOSE_ASSERT_EQ(false, doc->get_HyphenationOptions()->get_AutoHyphenation());
    ASSERT_EQ(0, doc->get_HyphenationOptions()->get_ConsecutiveHyphenLimit());
    ASSERT_EQ(360, doc->get_HyphenationOptions()->get_HyphenationZone());
    // 0.25 inch
    ASPOSE_ASSERT_EQ(true, doc->get_HyphenationOptions()->get_HyphenateCaps());
}

namespace gtest_test
{

TEST_F(ExDocument, HyphenationOptionsDefaultValues)
{
    s_instance->HyphenationOptionsDefaultValues();
}

} // namespace gtest_test

void ExDocument::HyphenationOptionsExceptions()
{
    auto doc = MakeObject<Document>();

    doc->get_HyphenationOptions()->set_ConsecutiveHyphenLimit(0);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_14 = [&doc]()
    {
        doc->get_HyphenationOptions()->set_HyphenationZone(0);
    };

    ASSERT_THROW(_local_func_14(), System::ArgumentOutOfRangeException);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_15 = [&doc]()
    {
        doc->get_HyphenationOptions()->set_ConsecutiveHyphenLimit(-1);
    };

    ASSERT_THROW(_local_func_15(), System::ArgumentOutOfRangeException);
    doc->get_HyphenationOptions()->set_HyphenationZone(360);
}

namespace gtest_test
{

TEST_F(ExDocument, HyphenationOptionsExceptions)
{
    s_instance->HyphenationOptionsExceptions();
}

} // namespace gtest_test

void ExDocument::ExtractPlainTextFromDocument()
{
    //ExStart
    //ExFor:PlainTextDocument
    //ExFor:PlainTextDocument.#ctor(String)
    //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
    //ExFor:PlainTextDocument.Text
    //ExSummary:Shows how to simply extract text from a document.
    auto loadOptions = MakeObject<TxtLoadOptions>();
    loadOptions->set_DetectNumberingWithWhitespaces(false);

    auto plaintext = MakeObject<PlainTextDocument>(MyDir + u"Document.docx");
    ASSERT_EQ(u"Hello World!", plaintext->get_Text().Trim());
    //ExSkip

    plaintext = MakeObject<PlainTextDocument>(MyDir + u"Document.docx", loadOptions);
    ASSERT_EQ(u"Hello World!", plaintext->get_Text().Trim());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ExtractPlainTextFromDocument)
{
    s_instance->ExtractPlainTextFromDocument();
}

} // namespace gtest_test

void ExDocument::GetPlainTextBuiltInDocumentProperties()
{
    //ExStart
    //ExFor:PlainTextDocument.BuiltInDocumentProperties
    //ExSummary:Shows how to get BuiltIn properties of plain text document.
    auto plaintext = MakeObject<PlainTextDocument>(MyDir + u"Bookmarks.docx");
    SharedPtr<BuiltInDocumentProperties> builtInDocumentProperties = plaintext->get_BuiltInDocumentProperties();
    //ExEnd

    ASSERT_EQ(u"Aspose", builtInDocumentProperties->get_Company());
}

namespace gtest_test
{

TEST_F(ExDocument, GetPlainTextBuiltInDocumentProperties)
{
    s_instance->GetPlainTextBuiltInDocumentProperties();
}

} // namespace gtest_test

void ExDocument::GetPlainTextCustomDocumentProperties()
{
    //ExStart
    //ExFor:PlainTextDocument.CustomDocumentProperties
    //ExSummary:Shows how to get custom properties of plain text document.
    auto plaintext = MakeObject<PlainTextDocument>(MyDir + u"Bookmarks.docx");
    SharedPtr<CustomDocumentProperties> customDocumentProperties = plaintext->get_CustomDocumentProperties();
    //ExEnd

    ASSERT_EQ(customDocumentProperties->get_Count(), 0);
}

namespace gtest_test
{

TEST_F(ExDocument, GetPlainTextCustomDocumentProperties)
{
    s_instance->GetPlainTextCustomDocumentProperties();
}

} // namespace gtest_test

void ExDocument::ExtractPlainTextFromStream()
{
    //ExStart
    //ExFor:PlainTextDocument.#ctor(Stream)
    //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
    //ExSummary:Shows how to simply extract text from a stream.
    auto loadOptions = MakeObject<TxtLoadOptions>();
    loadOptions->set_DetectNumberingWithWhitespaces(false);

    {
        SharedPtr<System::IO::Stream> stream = MakeObject<System::IO::FileStream>(MyDir + u"Document.docx", System::IO::FileMode::Open);
        auto plaintext = MakeObject<PlainTextDocument>(stream);
        ASSERT_EQ(u"Hello World!", plaintext->get_Text().Trim());
        //ExSkip

        plaintext = MakeObject<PlainTextDocument>(stream, loadOptions);
        ASSERT_EQ(u"Hello World!", plaintext->get_Text().Trim());
        //ExSkip
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, ExtractPlainTextFromStream)
{
    s_instance->ExtractPlainTextFromStream();
}

} // namespace gtest_test

void ExDocument::OoxmlComplianceVersion()
{
    //ExStart
    //ExFor:Document.Compliance
    //ExSummary:Shows how to get OOXML compliance version.
    // Open a DOC and check its OOXML compliance version
    auto doc = MakeObject<Document>(MyDir + u"Document.doc");

    OoxmlCompliance compliance = doc->get_Compliance();
    ASSERT_EQ(compliance, Aspose::Words::Saving::OoxmlCompliance::Ecma376_2006);

    // Open a DOCX which should have a newer one
    doc = MakeObject<Document>(MyDir + u"Document.docx");
    compliance = doc->get_Compliance();

    ASSERT_EQ(compliance, Aspose::Words::Saving::OoxmlCompliance::Iso29500_2008_Transitional);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, OoxmlComplianceVersion)
{
    s_instance->OoxmlComplianceVersion();
}

} // namespace gtest_test

void ExDocument::ImageSaveOptions()
{
    //ExStart
    //ExFor:Document.Save(String, Saving.SaveOptions)
    //ExFor:SaveOptions.UseAntiAliasing
    //ExFor:SaveOptions.UseHighQualityRendering
    //ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
    auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->get_Font()->set_Size(60);
    builder->Writeln(u"Some text.");

    SharedPtr<SaveOptions> options = MakeObject<Aspose::Words::Saving::ImageSaveOptions>(Aspose::Words::SaveFormat::Jpeg);
    ASSERT_FALSE(options->get_UseAntiAliasing());
    //ExSkip
    ASSERT_FALSE(options->get_UseHighQualityRendering());
    //ExSkip

    doc->Save(ArtifactsDir + u"Document.ImageSaveOptions.Default.jpg", options);

    options->set_UseAntiAliasing(true);
    options->set_UseHighQualityRendering(true);

    doc->Save(ArtifactsDir + u"Document.ImageSaveOptions.HighQuality.jpg", options);
    //ExEnd

    TestUtil::VerifyImage(794, 1122, ArtifactsDir + u"Document.ImageSaveOptions.Default.jpg");
    TestUtil::VerifyImage(794, 1122, ArtifactsDir + u"Document.ImageSaveOptions.HighQuality.jpg");
}

namespace gtest_test
{

TEST_F(ExDocument, DISABLED_ImageSaveOptions)
{
    s_instance->ImageSaveOptions();
}

} // namespace gtest_test

void ExDocument::CleanUpStyles()
{
    //ExStart
    //ExFor:Document.Cleanup()
    //ExSummary:Shows how to remove unused styles and lists from a document.
    // Create a new document
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add two styles and apply them to the builder's formats, marking them as "used"
    builder->get_ParagraphFormat()->set_Style(doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"My Used Style"));
    builder->get_ListFormat()->set_List(doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::BulletDiamonds));

    // And two more styles and leave them unused by not applying them to anything
    doc->get_Styles()->Add(Aspose::Words::StyleType::Paragraph, u"My Unused Style");
    doc->get_Lists()->Add(Aspose::Words::Lists::ListTemplate::NumberArabicDot);
    ASSERT_FALSE(doc->get_Styles()->idx_get(u"My Used Style") == nullptr);
    //ExSkip
    ASSERT_FALSE(doc->get_Styles()->idx_get(u"My Unused Style") == nullptr);
    //ExSkip

    auto hasBulletNumberStyle = [](SharedPtr<Lists::List> l)
    {
        return l->get_ListLevels()->idx_get(0)->get_NumberStyle() == NumberStyle::Bullet;
    };

    ASSERT_TRUE(doc->get_Lists()->LINQ_Any(hasBulletNumberStyle));
    //ExSkip

    auto hasArabicNumberStyle = [](SharedPtr<Lists::List> l)
    {
        return l->get_ListLevels()->idx_get(0)->get_NumberStyle() == NumberStyle::Arabic;
    };

    ASSERT_TRUE(doc->get_Lists()->LINQ_Any(hasArabicNumberStyle));
    //ExSkip

    doc->Cleanup();

    // The used styles are still in the document
    ASSERT_FALSE(doc->get_Styles()->idx_get(u"My Used Style") == nullptr);

    ASSERT_TRUE(doc->get_Lists()->LINQ_Any(hasBulletNumberStyle));

    // The unused styles have been removed
    ASSERT_TRUE(doc->get_Styles()->idx_get(u"My Unused Style") == nullptr);

    ASSERT_FALSE(doc->get_Lists()->LINQ_Any(hasArabicNumberStyle));
    //ExEnd

    ASSERT_EQ(5, doc->get_Styles()->get_Count());
    ASSERT_EQ(1, doc->get_Lists()->get_Count());

    doc->RemoveAllChildren();
    doc->Cleanup();

    ASSERT_EQ(4, doc->get_Styles()->get_Count());
    ASSERT_EQ(0, doc->get_Lists()->get_Count());
}

namespace gtest_test
{

TEST_F(ExDocument, CleanUpStyles)
{
    s_instance->CleanUpStyles();
}

} // namespace gtest_test

void ExDocument::Revisions()
{
    //ExStart
    //ExFor:Revision
    //ExFor:Revision.Accept
    //ExFor:Revision.Author
    //ExFor:Revision.DateTime
    //ExFor:Revision.Group
    //ExFor:Revision.Reject
    //ExFor:Revision.RevisionType
    //ExFor:RevisionCollection
    //ExFor:RevisionCollection.Item(Int32)
    //ExFor:RevisionCollection.Count
    //ExFor:RevisionType
    //ExFor:Document.HasRevisions
    //ExFor:Document.TrackRevisions
    //ExFor:Document.Revisions
    //ExSummary:Shows how to check if a document has revisions.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Normal editing of the document does not count as a revision
    builder->Write(u"This does not count as a revision. ");
    ASSERT_FALSE(doc->get_HasRevisions());

    // In order for our edits to count as revisions, we need to declare an author and start tracking them
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
    builder->Write(u"This is revision #1. ");

    // This flag corresponds to the "Track Changes" option being turned on in Microsoft Word, to track the editing manually
    // done there and not the programmatic changes we are about to do here
    ASSERT_FALSE(doc->get_TrackRevisions());

    // As well as nodes in the document, revisions get referenced in this collection
    ASSERT_TRUE(doc->get_HasRevisions());
    ASSERT_EQ(1, doc->get_Revisions()->get_Count());

    SharedPtr<Revision> revision = doc->get_Revisions()->idx_get(0);
    ASSERT_EQ(u"John Doe", revision->get_Author());
    ASSERT_EQ(u"This is revision #1. ", revision->get_ParentNode()->GetText());
    ASSERT_EQ(Aspose::Words::RevisionType::Insertion, revision->get_RevisionType());
    ASSERT_EQ(revision->get_DateTime().get_Date(), System::DateTime::get_Now().get_Date());
    ASPOSE_ASSERT_EQ(doc->get_Revisions()->get_Groups()->idx_get(0), revision->get_Group());

    // Deleting content also counts as a revision
    // The most recent revisions are put at the start of the collection
    doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->Remove();
    ASSERT_EQ(Aspose::Words::RevisionType::Deletion, doc->get_Revisions()->idx_get(0)->get_RevisionType());
    ASSERT_EQ(2, doc->get_Revisions()->get_Count());

    // Insert revisions are treated as document text by the GetText() method before they are accepted,
    // since they are still nodes with text and are in the body
    ASSERT_EQ(u"This does not count as a revision. This is revision #1.", doc->GetText().Trim());

    // Accepting the deletion revision will assimilate it into the paragraph text and remove it from the collection
    doc->get_Revisions()->idx_get(0)->Accept();
    ASSERT_EQ(1, doc->get_Revisions()->get_Count());

    // Once the delete revision is accepted, the nodes that it concerns are removed and their text will not show up here
    ASSERT_EQ(u"This is revision #1.", doc->GetText().Trim());

    // The second insertion revision is now at index 0, which we can reject to ignore and discard it
    doc->get_Revisions()->idx_get(0)->Reject();
    ASSERT_EQ(0, doc->get_Revisions()->get_Count());
    ASSERT_EQ(u"", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, Revisions)
{
    s_instance->Revisions();
}

} // namespace gtest_test

void ExDocument::RevisionCollection()
{
    //ExStart
    //ExFor:Revision.ParentStyle
    //ExFor:RevisionCollection.GetEnumerator
    //ExFor:RevisionCollection.Groups
    //ExFor:RevisionCollection.RejectAll
    //ExFor:RevisionGroupCollection.GetEnumerator
    //ExSummary:Shows how to look through a document's revisions.
    // Open a document that contains revisions and get its revision collection
    auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");
    SharedPtr<Aspose::Words::RevisionCollection> revisions = doc->get_Revisions();

    // This collection itself has a collection of revision groups, which are merged sequences of adjacent revisions
    ASSERT_EQ(7, revisions->get_Groups()->get_Count());
    //ExSkip
    System::Console::WriteLine(String::Format(u"{0} revision groups:",revisions->get_Groups()->get_Count()));

    // We can iterate over the collection of groups and access the text that the revision concerns
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<RevisionGroup>>> e = revisions->get_Groups()->GetEnumerator();
        while (e->MoveNext())
        {
            System::Console::WriteLine(String::Format(u"\tGroup type \"{0}\", ",e->get_Current()->get_RevisionType()) + String::Format(u"author: {0}, contents: [{1}]",e->get_Current()->get_Author(),e->get_Current()->get_Text().Trim()));
        }
    }

    // The collection of revisions is considerably larger than the condensed form we printed above,
    // depending on how many Runs the text has been segmented into during editing in Microsoft Word,
    // since each Run affected by a revision gets its own Revision object
    ASSERT_EQ(11, revisions->get_Count());
    //ExSkip
    System::Console::WriteLine(String::Format(u"\n{0} revisions:",revisions->get_Count()));

    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Revision>>> e = revisions->GetEnumerator();
        while (e->MoveNext())
        {
            // A StyleDefinitionChange strictly affects styles and not document nodes, so in this case the ParentStyle
            // attribute will always be used, while the ParentNode will always be null
            // Since all other changes affect nodes, ParentNode will conversely be in use and ParentStyle will be null
            if (e->get_Current()->get_RevisionType() == Aspose::Words::RevisionType::StyleDefinitionChange)
            {
                System::Console::WriteLine(String::Format(u"\tRevision type \"{0}\", ",e->get_Current()->get_RevisionType()) + String::Format(u"author: {0}, style: [{1}]",e->get_Current()->get_Author(),e->get_Current()->get_ParentStyle()->get_Name()));
            }
            else
            {
                System::Console::WriteLine(String::Format(u"\tRevision type \"{0}\", ",e->get_Current()->get_RevisionType()) + String::Format(u"author: {0}, contents: [{1}]",e->get_Current()->get_Author(),e->get_Current()->get_ParentNode()->GetText().Trim()));
            }
        }
    }

    // While the collection of revision groups provides a clearer overview of all revisions that took place in the document,
    // the changes must be accepted/rejected by the revisions themselves, the RevisionCollection, or the document
    // In this case we will reject all revisions via the collection, reverting the document to its original form, which we will then save
    revisions->RejectAll();
    ASSERT_EQ(0, revisions->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, RevisionCollection)
{
    s_instance->RevisionCollection();
}

} // namespace gtest_test

void ExDocument::AutomaticallyUpdateStyles()
{
    //ExStart
    //ExFor:Document.AutomaticallyUpdateStyles
    //ExSummary:Shows how to update a document's styles based on its template.
    auto doc = MakeObject<Document>();

    // Empty Microsoft Word documents by default come with an attached template called "Normal.dotm"
    // There is no default template for Aspose Words documents
    ASSERT_EQ(String::Empty, doc->get_AttachedTemplate());

    // For AutomaticallyUpdateStyles to have any effect, we need a document with a template
    // We can make a document with word and open it
    // Or we can attach a template from our file system, as below
    doc->set_AttachedTemplate(MyDir + u"Business brochure.dotx");

    // Any changes to the styles in this template will be propagated to those styles in the document
    doc->set_AutomaticallyUpdateStyles(true);

    doc->Save(ArtifactsDir + u"Document.AutomaticallyUpdateStyles.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.AutomaticallyUpdateStyles.docx");

    ASSERT_TRUE(doc->get_AutomaticallyUpdateStyles());
    ASSERT_EQ(MyDir + u"Business brochure.dotx", doc->get_AttachedTemplate());
    ASSERT_TRUE(System::IO::File::Exists(doc->get_AttachedTemplate()));
}

namespace gtest_test
{

TEST_F(ExDocument, AutomaticallyUpdateStyles)
{
    s_instance->AutomaticallyUpdateStyles();
}

} // namespace gtest_test

void ExDocument::DefaultTemplate()
{
    //ExStart
    //ExFor:Document.AttachedTemplate
    //ExFor:SaveOptions.CreateSaveOptions(String)
    //ExFor:SaveOptions.DefaultTemplate
    //ExSummary:Shows how to set a default .docx document template.
    auto doc = MakeObject<Document>();

    // If we set this flag to true while not having a template attached to the document,
    // there will be no effect because there is no template document to draw style changes from
    doc->set_AutomaticallyUpdateStyles(true);
    ASSERT_EQ(doc->get_AttachedTemplate().get_Length(), 0);

    // We can set a default template document filename in a SaveOptions object to make it apply to
    // all documents we save with it that have no AttachedTemplate value
    SharedPtr<SaveOptions> options = SaveOptions::CreateSaveOptions(u"Document.DefaultTemplate.docx");
    options->set_DefaultTemplate(MyDir + u"Business brochure.dotx");
    ASSERT_TRUE(System::IO::File::Exists(options->get_DefaultTemplate()));
    //ExSkip

    doc->Save(ArtifactsDir + u"Document.DefaultTemplate.docx", options);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DefaultTemplate)
{
    s_instance->DefaultTemplate();
}

} // namespace gtest_test

void ExDocument::Sections()
{
    //ExStart
    //ExFor:Document.LastSection
    //ExSummary:Shows how to edit the last section of a document.
    // Open the template document, containing obsolete copyright information in the footer
    auto doc = MakeObject<Document>(MyDir + u"Footer.docx");

    // Create a new copyright information string to replace an older one with
    int currentYear = System::DateTime::get_Now().get_Year();
    String newCopyrightInformation = String::Format(u"Copyright (C) {0} by Aspose Pty Ltd.",currentYear);

    auto findReplaceOptions = MakeObject<FindReplaceOptions>();
    findReplaceOptions->set_MatchCase(false);
    findReplaceOptions->set_FindWholeWordsOnly(false);

    // Each section has its own set of headers/footers,
    // so the text in each one has to be replaced individually if we want the entire document to be affected
    SharedPtr<HeaderFooter> firstSectionFooter = doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary);
    firstSectionFooter->get_Range()->Replace(u"(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

    SharedPtr<HeaderFooter> lastSectionFooter = doc->get_LastSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary);
    lastSectionFooter->get_Range()->Replace(u"(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

    doc->Save(ArtifactsDir + u"Document.Sections.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.Sections.docx");
    ASPOSE_ASSERT_EQ(doc->get_FirstSection(), doc->get_Sections()->idx_get(0));
    ASPOSE_ASSERT_EQ(doc->get_LastSection(), doc->get_Sections()->idx_get(1));

    ASSERT_TRUE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Contains(String::Format(u"Copyright (C) {0} by Aspose Pty Ltd.",currentYear)));
    ASSERT_TRUE(doc->get_LastSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Contains(String::Format(u"Copyright (C) {0} by Aspose Pty Ltd.",currentYear)));
    ASSERT_FALSE(doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Contains(u"(C) 2006 Aspose Pty Ltd."));
    ASSERT_FALSE(doc->get_LastSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Contains(u"(C) 2006 Aspose Pty Ltd."));
}

namespace gtest_test
{

TEST_F(ExDocument, Sections)
{
    s_instance->Sections();
}

} // namespace gtest_test

void ExDocument::UseLegacyOrder(bool isUseLegacyOrder)
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert 3 tags to appear in sequential order, the second of which will be inside a text box
    builder->Writeln(u"[tag 1]");
    SharedPtr<Shape> textBox = builder->InsertShape(Aspose::Words::Drawing::ShapeType::TextBox, 100, 50);
    builder->Writeln(u"[tag 3]");

    builder->MoveTo(textBox->get_FirstParagraph());
    builder->Write(u"[tag 2]");

    auto callback = MakeObject<ExDocument::UseLegacyOrderReplacingCallback>();
    auto options = MakeObject<FindReplaceOptions>();
    options->set_ReplacingCallback(callback);

    // Use this option if want to search text sequentially from top to bottom considering the text boxes
    options->set_UseLegacyOrder(isUseLegacyOrder);

    doc->get_Range()->Replace(MakeObject<System::Text::RegularExpressions::Regex>(u"\\[(.*?)\\]"), u"", options);
    CheckUseLegacyOrderResults(isUseLegacyOrder, callback);
    //ExSkip

    for (auto match : (System::StaticCast<ApiExamples::ExDocument::UseLegacyOrderReplacingCallback>(options->get_ReplacingCallback()))->Matches)
    {
        System::Console::WriteLine(match);
    }
}

namespace gtest_test
{

using UseLegacyOrder_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExDocument::UseLegacyOrder)>::type;

struct UseLegacyOrderVP : public ExDocument, public ApiExamples::ExDocument, public ::testing::WithParamInterface<UseLegacyOrder_Args>
{
    static std::vector<UseLegacyOrder_Args> TestCases()
    {
        return
        {
            std::make_tuple(true),
            std::make_tuple(false),
        };
    }
};

TEST_P(UseLegacyOrderVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UseLegacyOrder(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExDocument, UseLegacyOrderVP, ::testing::ValuesIn(UseLegacyOrderVP::TestCases()));

} // namespace gtest_test

void ExDocument::SetInvalidateFieldTypes()
{
    //ExStart
    //ExFor:Document.NormalizeFieldTypes
    //ExSummary:Shows how to get the keep a field's type up to date with its field code.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add a date field
    SharedPtr<Field> field = builder->InsertField(u"DATE", nullptr);

    // Based on the field code we entered above, the type of the field has been set to "FieldDate"
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Type());

    // We can manually access the content of the field we added and change it
    auto fieldText = System::DynamicCast<Aspose::Words::Run>(doc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetChildNodes(Aspose::Words::NodeType::Run, true)->idx_get(0));
    ASSERT_EQ(u"DATE", fieldText->get_Text());
    //ExSkip
    fieldText->set_Text(u"PAGE");

    // We changed the text to "PAGE" but the field's type property did not update accordingly
    ASSERT_EQ(u"PAGE", fieldText->GetText());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Type());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Start()->get_FieldType());
    //ExSkip
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_Separator()->get_FieldType());
    //ExSkip
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldDate, field->get_End()->get_FieldType());
    //ExSkip

    // After running this method the type of the field, as well as its FieldStart,
    // FieldSeparator and FieldEnd nodes to "FieldPage", which matches the text "PAGE"
    doc->NormalizeFieldTypes();

    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_Type());
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_Start()->get_FieldType());
    //ExSkip
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_Separator()->get_FieldType());
    //ExSkip
    ASSERT_EQ(Aspose::Words::Fields::FieldType::FieldPage, field->get_End()->get_FieldType());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SetInvalidateFieldTypes)
{
    s_instance->SetInvalidateFieldTypes();
}

} // namespace gtest_test

void ExDocument::LayoutOptions()
{
    //ExStart
    //ExFor:Document.LayoutOptions
    //ExFor:LayoutOptions
    //ExFor:LayoutOptions.RevisionOptions
    //ExFor:Layout.LayoutOptions.ShowHiddenText
    //ExFor:Layout.LayoutOptions.ShowParagraphMarks
    //ExFor:RevisionColor
    //ExFor:RevisionOptions
    //ExFor:RevisionOptions.InsertedTextColor
    //ExFor:RevisionOptions.ShowRevisionBars
    //ExSummary:Shows how to set a document's layout options.
    auto doc = MakeObject<Document>();
    SharedPtr<Aspose::Words::Layout::LayoutOptions> options = doc->get_LayoutOptions();
    ASSERT_FALSE(options->get_ShowHiddenText());
    //ExSkip
    ASSERT_FALSE(options->get_ShowParagraphMarks());
    //ExSkip

    // The appearance of revisions can be controlled from the layout options property
    doc->StartTrackRevisions(u"John Doe", System::DateTime::get_Now());
    ASSERT_EQ(Aspose::Words::Layout::RevisionColor::ByAuthor, options->get_RevisionOptions()->get_InsertedTextColor());
    //ExSkip
    ASSERT_TRUE(options->get_RevisionOptions()->get_ShowRevisionBars());
    //ExSkip
    options->get_RevisionOptions()->set_InsertedTextColor(Aspose::Words::Layout::RevisionColor::BrightGreen);
    options->get_RevisionOptions()->set_ShowRevisionBars(false);

    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"This is a revision. Normally the text is red with a bar to the left, but we made some changes to the revision options.");

    doc->StopTrackRevisions();

    // Layout options can be used to show hidden text too
    builder->Writeln(u"This text is not hidden.");
    builder->get_Font()->set_Hidden(true);
    builder->Writeln(u"This text is hidden. It will only show up in the output if we allow it to via doc.LayoutOptions.");

    options->set_ShowHiddenText(true);

    // This option is equivalent to enabling paragraph marks in Microsoft Word via Home > paragraph > Show Paragraph Marks,
    // and can be used to display these features in a .pdf
    options->set_ShowParagraphMarks(true);

    doc->Save(ArtifactsDir + u"Document.LayoutOptions.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, LayoutOptions)
{
    s_instance->LayoutOptions();
}

} // namespace gtest_test

void ExDocument::MailMergeSettings()
{
    //ExStart
    //ExFor:Document.MailMergeSettings
    //ExFor:MailMergeCheckErrors
    //ExFor:MailMergeDataType
    //ExFor:MailMergeDestination
    //ExFor:MailMergeMainDocumentType
    //ExFor:MailMergeSettings
    //ExFor:MailMergeSettings.CheckErrors
    //ExFor:MailMergeSettings.Clone
    //ExFor:MailMergeSettings.Destination
    //ExFor:MailMergeSettings.DataType
    //ExFor:MailMergeSettings.DoNotSupressBlankLines
    //ExFor:MailMergeSettings.LinkToQuery
    //ExFor:MailMergeSettings.MainDocumentType
    //ExFor:MailMergeSettings.Odso
    //ExFor:MailMergeSettings.Query
    //ExFor:MailMergeSettings.ViewMergedData
    //ExFor:Odso
    //ExFor:Odso.Clone
    //ExFor:Odso.ColumnDelimiter
    //ExFor:Odso.DataSource
    //ExFor:Odso.DataSourceType
    //ExFor:Odso.FirstRowContainsColumnNames
    //ExFor:OdsoDataSourceType
    //ExSummary:Shows how to execute an Office Data Source Object mail merge with MailMergeSettings.
    // We'll create a simple document that will act as a destination for mail merge data
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Dear ");
    builder->InsertField(u"MERGEFIELD FirstName", u"<FirstName>");
    builder->Write(u" ");
    builder->InsertField(u"MERGEFIELD LastName", u"<LastName>");
    builder->Writeln(u": ");
    builder->InsertField(u"MERGEFIELD Message", u"<Message>");

    // Also we'll need a data source, in this case it will be an ASCII text file
    // We can use any character we want as a delimiter, in this case we'll choose '|'
    // The delimiter character is selected in the ODSO settings of mail merge settings
    ArrayPtr<String> lines = MakeArray<String>({u"FirstName|LastName|Message", u"John|Doe|Hello! This message was created with Aspose Words mail merge."});
    String dataSrcFilename = ArtifactsDir + u"Document.MailMergeSettings.DataSource.txt";

    System::IO::File::WriteAllLines(dataSrcFilename, lines);

    // Set the data source, query and other things
    SharedPtr<Aspose::Words::Settings::MailMergeSettings> settings = doc->get_MailMergeSettings();
    settings->set_MainDocumentType(Aspose::Words::Settings::MailMergeMainDocumentType::MailingLabels);
    settings->set_CheckErrors(Aspose::Words::Settings::MailMergeCheckErrors::Simulate);
    settings->set_DataType(Aspose::Words::Settings::MailMergeDataType::Native);
    settings->set_DataSource(dataSrcFilename);
    settings->set_Query(String(u"SELECT * FROM ") + doc->get_MailMergeSettings()->get_DataSource());
    settings->set_LinkToQuery(true);
    settings->set_ViewMergedData(true);

    ASSERT_EQ(Aspose::Words::Settings::MailMergeDestination::Default, settings->get_Destination());
    ASSERT_FALSE(settings->get_DoNotSupressBlankLines());

    // Office Data Source Object settings
    SharedPtr<Odso> odso = settings->get_Odso();
    odso->set_DataSource(dataSrcFilename);
    odso->set_DataSourceType(Aspose::Words::Settings::OdsoDataSourceType::Text);
    odso->set_ColumnDelimiter(u'|');
    odso->set_FirstRowContainsColumnNames(true);

    // ODSO/MailMergeSettings objects can also be cloned
    ASPOSE_ASSERT_NS(odso, odso->Clone());
    ASPOSE_ASSERT_NS(settings, settings->Clone());

    // The mail merge will be performed when this document is opened
    doc->Save(ArtifactsDir + u"Document.MailMergeSettings.docx");
    //ExEnd

    settings = MakeObject<Document>(ArtifactsDir + u"Document.MailMergeSettings.docx")->get_MailMergeSettings();

    ASSERT_EQ(Aspose::Words::Settings::MailMergeMainDocumentType::MailingLabels, settings->get_MainDocumentType());
    ASSERT_EQ(Aspose::Words::Settings::MailMergeCheckErrors::Simulate, settings->get_CheckErrors());
    ASSERT_EQ(Aspose::Words::Settings::MailMergeDataType::Native, settings->get_DataType());
    ASSERT_EQ(ArtifactsDir + u"Document.MailMergeSettings.DataSource.txt", settings->get_DataSource());
    ASSERT_EQ(String(u"SELECT * FROM ") + doc->get_MailMergeSettings()->get_DataSource(), settings->get_Query());
    ASSERT_TRUE(settings->get_LinkToQuery());
    ASSERT_TRUE(settings->get_ViewMergedData());

    odso = settings->get_Odso();
    ASSERT_EQ(ArtifactsDir + u"Document.MailMergeSettings.DataSource.txt", odso->get_DataSource());
    ASSERT_EQ(Aspose::Words::Settings::OdsoDataSourceType::Text, odso->get_DataSourceType());
    ASPOSE_ASSERT_EQ(u'|', odso->get_ColumnDelimiter());
    ASSERT_TRUE(odso->get_FirstRowContainsColumnNames());
}

namespace gtest_test
{

TEST_F(ExDocument, MailMergeSettings)
{
    s_instance->MailMergeSettings();
}

} // namespace gtest_test

void ExDocument::OdsoEmail()
{
    //ExStart
    //ExFor:MailMergeSettings.ActiveRecord
    //ExFor:MailMergeSettings.AddressFieldName
    //ExFor:MailMergeSettings.ConnectString
    //ExFor:MailMergeSettings.MailAsAttachment
    //ExFor:MailMergeSettings.MailSubject
    //ExFor:MailMergeSettings.Clear
    //ExFor:Odso.TableName
    //ExFor:Odso.UdlConnectString
    //ExSummary:Shows how to execute a mail merge while connecting to an external data source.
    auto doc = MakeObject<Document>(MyDir + u"Odso data.docx");
    TestOdsoEmail(doc);
    //ExSkip
    SharedPtr<Aspose::Words::Settings::MailMergeSettings> settings = doc->get_MailMergeSettings();

    System::Console::WriteLine(String::Format(u"Connection string:\n\t{0}",settings->get_ConnectString()));
    System::Console::WriteLine(String::Format(u"Mail merge docs as attachment:\n\t{0}",settings->get_MailAsAttachment()));
    System::Console::WriteLine(String::Format(u"Mail merge doc e-mail subject:\n\t{0}",settings->get_MailSubject()));
    System::Console::WriteLine(String::Format(u"Column that contains e-mail addresses:\n\t{0}",settings->get_AddressFieldName()));
    System::Console::WriteLine(String::Format(u"Active record:\n\t{0}",settings->get_ActiveRecord()));

    SharedPtr<Odso> odso = settings->get_Odso();

    System::Console::WriteLine(String::Format(u"File will connect to data source located in:\n\t\"{0}\"",odso->get_DataSource()));
    System::Console::WriteLine(String::Format(u"Source type:\n\t{0}",odso->get_DataSourceType()));
    System::Console::WriteLine(String::Format(u"UDL connection string:\n\t{0}",odso->get_UdlConnectString()));
    System::Console::WriteLine(String::Format(u"Table:\n\t{0}",odso->get_TableName()));
    System::Console::WriteLine(String::Format(u"Query:\n\t{0}",doc->get_MailMergeSettings()->get_Query()));

    // We can clear the settings, which will take place during saving
    settings->Clear();

    doc->Save(ArtifactsDir + u"Document.OdsoEmail.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.OdsoEmail.docx");
    ASSERT_TRUE(String::IsNullOrEmpty(doc->get_MailMergeSettings()->get_ConnectString()));
}

namespace gtest_test
{

TEST_F(ExDocument, OdsoEmail)
{
    s_instance->OdsoEmail();
}

} // namespace gtest_test

void ExDocument::MailingLabelMerge()
{
    //ExStart
    //ExFor:MailMergeSettings.DataSource
    //ExFor:MailMergeSettings.HeaderSource
    //ExSummary:Shows how to execute a mail merge while drawing data from a header and a data file.
    // Create a mailing label merge header file, which will consist of a table with one row
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"FirstName");
    builder->InsertCell();
    builder->Write(u"LastName");
    builder->EndTable();

    doc->Save(ArtifactsDir + u"Document.MailingLabelMerge.Header.docx");

    // Create a mailing label merge date file, which will consist of a table with one row and the same amount of columns as
    // the header table, which will determine the names for these columns
    doc = MakeObject<Document>();
    builder = MakeObject<DocumentBuilder>(doc);

    builder->StartTable();
    builder->InsertCell();
    builder->Write(u"John");
    builder->InsertCell();
    builder->Write(u"Doe");
    builder->EndTable();

    doc->Save(ArtifactsDir + u"Document.MailingLabelMerge.Data.docx");

    // Create a merge destination document with MERGEFIELDS that will accept data
    doc = MakeObject<Document>();
    builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Dear ");
    builder->InsertField(u"MERGEFIELD FirstName", u"<FirstName>");
    builder->Write(u" ");
    builder->InsertField(u"MERGEFIELD LastName", u"<LastName>");

    // Configure settings to draw data and headers from other documents
    SharedPtr<Aspose::Words::Settings::MailMergeSettings> settings = doc->get_MailMergeSettings();

    // The "header" document contains column names for the data in the "data" document,
    // which will correspond to the names of our MERGEFIELDs
    settings->set_HeaderSource(ArtifactsDir + u"Document.MailingLabelMerge.Header.docx");
    settings->set_DataSource(ArtifactsDir + u"Document.MailingLabelMerge.Data.docx");

    // Configure the rest of the MailMergeSettings object
    settings->set_Query(String(u"SELECT * FROM ") + settings->get_DataSource());
    settings->set_MainDocumentType(Aspose::Words::Settings::MailMergeMainDocumentType::MailingLabels);
    settings->set_DataType(Aspose::Words::Settings::MailMergeDataType::TextFile);
    settings->set_LinkToQuery(true);
    settings->set_ViewMergedData(true);

    // The mail merge will be performed when this document is opened
    doc->Save(ArtifactsDir + u"Document.MailingLabelMerge.docx");
    //ExEnd

    ASSERT_EQ(u"FirstName\aLastName\a\a", MakeObject<Document>(ArtifactsDir + u"Document.MailingLabelMerge.Header.docx")->GetChild(Aspose::Words::NodeType::Table, 0, true)->GetText().Trim());

    ASSERT_EQ(u"John\aDoe\a\a", MakeObject<Document>(ArtifactsDir + u"Document.MailingLabelMerge.Data.docx")->GetChild(Aspose::Words::NodeType::Table, 0, true)->GetText().Trim());

    doc = MakeObject<Document>(ArtifactsDir + u"Document.MailingLabelMerge.docx");

    ASSERT_EQ(2, doc->get_Range()->get_Fields()->get_Count());

    settings = doc->get_MailMergeSettings();

    ASSERT_EQ(ArtifactsDir + u"Document.MailingLabelMerge.Header.docx", settings->get_HeaderSource());
    ASSERT_EQ(ArtifactsDir + u"Document.MailingLabelMerge.Data.docx", settings->get_DataSource());
    ASSERT_EQ(String(u"SELECT * FROM ") + settings->get_DataSource(), settings->get_Query());
    ASSERT_EQ(Aspose::Words::Settings::MailMergeMainDocumentType::MailingLabels, settings->get_MainDocumentType());
    ASSERT_EQ(Aspose::Words::Settings::MailMergeDataType::TextFile, settings->get_DataType());
    ASSERT_TRUE(settings->get_LinkToQuery());
    ASSERT_TRUE(settings->get_ViewMergedData());
}

namespace gtest_test
{

TEST_F(ExDocument, MailingLabelMerge)
{
    s_instance->MailingLabelMerge();
}

} // namespace gtest_test

void ExDocument::OdsoFieldMapDataCollection()
{
    //ExStart
    //ExFor:Odso.FieldMapDatas
    //ExFor:OdsoFieldMapData
    //ExFor:OdsoFieldMapData.Clone
    //ExFor:OdsoFieldMapData.Column
    //ExFor:OdsoFieldMapData.MappedName
    //ExFor:OdsoFieldMapData.Name
    //ExFor:OdsoFieldMapData.Type
    //ExFor:OdsoFieldMapDataCollection
    //ExFor:OdsoFieldMapDataCollection.Add(OdsoFieldMapData)
    //ExFor:OdsoFieldMapDataCollection.Clear
    //ExFor:OdsoFieldMapDataCollection.Count
    //ExFor:OdsoFieldMapDataCollection.GetEnumerator
    //ExFor:OdsoFieldMapDataCollection.Item(Int32)
    //ExFor:OdsoFieldMapDataCollection.RemoveAt(Int32)
    //ExFor:OdsoFieldMappingType
    //ExSummary:Shows how to access the collection of data that maps data source columns to merge fields.
    auto doc = MakeObject<Document>(MyDir + u"Odso data.docx");

    // This collection defines how columns from an external data source will be mapped to predefined MERGEFIELD,
    // ADDRESSBLOCK and GREETINGLINE fields during a mail merge
    SharedPtr<Aspose::Words::Settings::OdsoFieldMapDataCollection> dataCollection = doc->get_MailMergeSettings()->get_Odso()->get_FieldMapDatas();
    ASSERT_EQ(30, dataCollection->get_Count());

    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<OdsoFieldMapData>>> enumerator = dataCollection->GetEnumerator();
        int index = 0;
        while (enumerator->MoveNext())
        {
            System::Console::WriteLine(String::Format(u"Field map data index {0}, type \"{1}\":",index++,enumerator->get_Current()->get_Type()));

            System::Console::WriteLine(enumerator->get_Current()->get_Type() != Aspose::Words::Settings::OdsoFieldMappingType::Null ? String::Format(u"\tColumn \"{0}\", number {1} mapped to merge field \"{2}\".",enumerator->get_Current()->get_Name(),enumerator->get_Current()->get_Column(),enumerator->get_Current()->get_MappedName()) : u"\tNo valid column to field mapping data present.");
        }
    }

    // Elements of the collection can be cloned
    ASPOSE_ASSERT_NE(dataCollection->idx_get(0), dataCollection->idx_get(0)->Clone());

    // The collection can have individual entries removed or be cleared like this
    dataCollection->RemoveAt(0);
    ASSERT_EQ(29, dataCollection->get_Count());
    //ExSkip
    dataCollection->Clear();
    ASSERT_EQ(0, dataCollection->get_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, OdsoFieldMapDataCollection)
{
    s_instance->OdsoFieldMapDataCollection();
}

} // namespace gtest_test

void ExDocument::OdsoRecipientDataCollection()
{
    //ExStart
    //ExFor:Odso.RecipientDatas
    //ExFor:OdsoRecipientData
    //ExFor:OdsoRecipientData.Active
    //ExFor:OdsoRecipientData.Clone
    //ExFor:OdsoRecipientData.Column
    //ExFor:OdsoRecipientData.Hash
    //ExFor:OdsoRecipientData.UniqueTag
    //ExFor:OdsoRecipientDataCollection
    //ExFor:OdsoRecipientDataCollection.Add(OdsoRecipientData)
    //ExFor:OdsoRecipientDataCollection.Clear
    //ExFor:OdsoRecipientDataCollection.Count
    //ExFor:OdsoRecipientDataCollection.GetEnumerator
    //ExFor:OdsoRecipientDataCollection.Item(Int32)
    //ExFor:OdsoRecipientDataCollection.RemoveAt(Int32)
    //ExSummary:Shows how to access the collection of data that designates merge data source records to be excluded from a merge.
    auto doc = MakeObject<Document>(MyDir + u"Odso data.docx");

    // Records in this collection that do not have the "Active" flag set to true will be excluded from the mail merge
    SharedPtr<Aspose::Words::Settings::OdsoRecipientDataCollection> dataCollection = doc->get_MailMergeSettings()->get_Odso()->get_RecipientDatas();

    ASSERT_EQ(70, dataCollection->get_Count());

    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<OdsoRecipientData>>> enumerator = dataCollection->GetEnumerator();
        int index = 0;
        while (enumerator->MoveNext())
        {
            System::Console::WriteLine(String::Format(u"Odso recipient data index {0} will {1}be imported upon mail merge.",index++,(enumerator->get_Current()->get_Active() ? String(u"") : String(u"not "))));
            System::Console::WriteLine(String::Format(u"\tColumn #{0}",enumerator->get_Current()->get_Column()));
            System::Console::WriteLine(String::Format(u"\tHash code: {0}",enumerator->get_Current()->get_Hash()));
            System::Console::WriteLine(String::Format(u"\tContents array length: {0}",enumerator->get_Current()->get_UniqueTag()->get_Length()));
        }
    }

    // Elements of the collection can be cloned
    ASPOSE_ASSERT_NE(dataCollection->idx_get(0), dataCollection->idx_get(0)->Clone());

    // The collection can have individual entries removed or be cleared like this
    dataCollection->RemoveAt(0);
    ASSERT_EQ(69, dataCollection->get_Count());
    //ExSkip
    dataCollection->Clear();
    ASSERT_EQ(0, dataCollection->get_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, OdsoRecipientDataCollection)
{
    s_instance->OdsoRecipientDataCollection();
}

} // namespace gtest_test

void ExDocument::DocPackageCustomParts()
{
    //ExStart
    //ExFor:CustomPart
    //ExFor:CustomPart.ContentType
    //ExFor:CustomPart.RelationshipType
    //ExFor:CustomPart.IsExternal
    //ExFor:CustomPart.Data
    //ExFor:CustomPart.Name
    //ExFor:CustomPart.Clone
    //ExFor:CustomPartCollection
    //ExFor:CustomPartCollection.Add(CustomPart)
    //ExFor:CustomPartCollection.Clear
    //ExFor:CustomPartCollection.Clone
    //ExFor:CustomPartCollection.Count
    //ExFor:CustomPartCollection.GetEnumerator
    //ExFor:CustomPartCollection.Item(Int32)
    //ExFor:CustomPartCollection.RemoveAt(Int32)
    //ExFor:Document.PackageCustomParts
    //ExSummary:Shows how to open a document with custom parts and access them.
    // Open a document that contains custom parts
    // CustomParts are arbitrary content OOXML parts
    // Not to be confused with Custom XML data which is represented by CustomXmlParts
    // This part is internal, meaning it is contained inside the OOXML package
    auto doc = MakeObject<Document>(MyDir + u"Custom parts OOXML package.docx");

    // Clone the second part
    SharedPtr<CustomPart> clonedPart = doc->get_PackageCustomParts()->idx_get(1)->Clone();

    // Add the clone to the collection
    doc->get_PackageCustomParts()->Add(clonedPart);
    TestDocPackageCustomParts(doc->get_PackageCustomParts());
    //ExSkip

    // Use an enumerator to print information about the contents of each part
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<CustomPart>>> enumerator = doc->get_PackageCustomParts()->GetEnumerator();
        int index = 0;
        while (enumerator->MoveNext())
        {
            System::Console::WriteLine(String::Format(u"Part index {0}:",index));
            System::Console::WriteLine(String::Format(u"\tName: {0}",enumerator->get_Current()->get_Name()));
            System::Console::WriteLine(String::Format(u"\tContentType: {0}",enumerator->get_Current()->get_ContentType()));
            System::Console::WriteLine(String::Format(u"\tRelationshipType: {0}",enumerator->get_Current()->get_RelationshipType()));
            System::Console::WriteLine(enumerator->get_Current()->get_IsExternal() ? u"\tSourced from outside the document" : String::Format(u"\tSourced from within the document, length: {0} bytes",enumerator->get_Current()->get_Data()->get_Length()));
            index++;
        }
    }

    // The parts collection can have individual entries removed or be cleared like this
    doc->get_PackageCustomParts()->RemoveAt(2);
    ASSERT_EQ(2, doc->get_PackageCustomParts()->get_Count());
    //ExSkip
    doc->get_PackageCustomParts()->Clear();
    ASSERT_EQ(0, doc->get_PackageCustomParts()->get_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, DocPackageCustomParts)
{
    s_instance->DocPackageCustomParts();
}

} // namespace gtest_test

void ExDocument::ShadeFormData()
{
    //ExStart
    //ExFor:Document.ShadeFormData
    //ExSummary:Shows how to apply gray shading to bookmarks.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // By default, bookmarked text is highlighted gray
    ASSERT_TRUE(doc->get_ShadeFormData());

    builder->Write(u"Text before bookmark. ");

    builder->InsertTextInput(u"My bookmark", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"If gray form field shading is turned on, this is the text that will have a gray background.", 0);

    // We can turn the grey shading off so the bookmarked text will blend in with the other text
    doc->set_ShadeFormData(false);
    doc->Save(ArtifactsDir + u"Document.ShadeFormData.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.ShadeFormData.docx");
    ASSERT_FALSE(doc->get_ShadeFormData());
}

namespace gtest_test
{

TEST_F(ExDocument, ShadeFormData)
{
    s_instance->ShadeFormData();
}

} // namespace gtest_test

void ExDocument::VersionsCount()
{
    //ExStart
    //ExFor:Document.VersionsCount
    //ExSummary:Shows how to count how many previous versions a document has.
    // Document versions are not supported but we can open an older document that has them
    auto doc = MakeObject<Document>(MyDir + u"Versions.doc");

    // We can use this property to see how many there are
    // If we save and open the document, they will be lost
    ASSERT_EQ(4, doc->get_VersionsCount());
    //ExEnd

    doc->Save(ArtifactsDir + u"Document.VersionsCount.docx");
    doc = MakeObject<Document>(ArtifactsDir + u"Document.VersionsCount.docx");

    ASSERT_EQ(0, doc->get_VersionsCount());
}

namespace gtest_test
{

TEST_F(ExDocument, VersionsCount)
{
    s_instance->VersionsCount();
}

} // namespace gtest_test

void ExDocument::WriteProtection()
{
    //ExStart
    //ExFor:Document.WriteProtection
    //ExFor:WriteProtection
    //ExFor:WriteProtection.IsWriteProtected
    //ExFor:WriteProtection.ReadOnlyRecommended
    //ExFor:WriteProtection.SetPassword(String)
    //ExFor:WriteProtection.ValidatePassword(String)
    //ExSummary:Shows how to protect a document with a password.
    auto doc = MakeObject<Document>();
    ASSERT_FALSE(doc->get_WriteProtection()->get_IsWriteProtected());
    //ExSkip
    ASSERT_FALSE(doc->get_WriteProtection()->get_ReadOnlyRecommended());
    //ExSkip

    // Enter a password that's up to 15 characters long
    doc->get_WriteProtection()->SetPassword(u"MyPassword");

    ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());
    ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));

    // This flag applies to RTF documents and will be ignored by Microsoft Word
    doc->get_WriteProtection()->set_ReadOnlyRecommended(true);

    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Write protection does not prevent us from editing the document programmatically.");

    // Save the document
    // Without the password, we can only read this document in Microsoft Word
    // With the password, we can read and write
    doc->Save(ArtifactsDir + u"Document.WriteProtection.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.WriteProtection.docx");

    ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());
    ASSERT_TRUE(doc->get_WriteProtection()->get_ReadOnlyRecommended());
    ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));
    ASSERT_FALSE(doc->get_WriteProtection()->ValidatePassword(u"wrongpassword"));

    builder = MakeObject<DocumentBuilder>(doc);
    builder->MoveToDocumentEnd();
    builder->Writeln(u"Writing text in a protected document.");

    ASSERT_EQ(String(u"Write protection does not prevent us from editing the document programmatically.") + u"\rWriting text in a protected document.", doc->GetText().Trim());
}

namespace gtest_test
{

TEST_F(ExDocument, WriteProtection)
{
    s_instance->WriteProtection();
}

} // namespace gtest_test

void ExDocument::AddEditingLanguage()
{
    //ExStart
    //ExFor:LanguagePreferences
    //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
    //ExFor:LoadOptions.LanguagePreferences
    //ExFor:EditingLanguage
    //ExSummary:Shows how to set up language preferences that will be used when document is loading
    auto loadOptions = MakeObject<LoadOptions>();
    loadOptions->get_LanguagePreferences()->AddEditingLanguage(Aspose::Words::EditingLanguage::Japanese);

    auto doc = MakeObject<Document>(MyDir + u"No default editing language.docx", loadOptions);

    int localeIdFarEast = doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast();
    System::Console::WriteLine(localeIdFarEast == (int)Aspose::Words::EditingLanguage::Japanese ? String(u"The document either has no any FarEast language set in defaults or it was set to Japanese originally.") : String(u"The document default FarEast language was set to another than Japanese language originally, so it is not overridden."));
    //ExEnd

    ASSERT_EQ((int)Aspose::Words::EditingLanguage::Japanese, doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast());

    doc = MakeObject<Document>(MyDir + u"No default editing language.docx");

    ASSERT_EQ((int)Aspose::Words::EditingLanguage::EnglishUS, doc->get_Styles()->get_DefaultFont()->get_LocaleIdFarEast());
}

namespace gtest_test
{

TEST_F(ExDocument, AddEditingLanguage)
{
    s_instance->AddEditingLanguage();
}

} // namespace gtest_test

void ExDocument::SetEditingLanguageAsDefault()
{
    //ExStart
    //ExFor:LanguagePreferences.DefaultEditingLanguage
    //ExSummary:Shows how to set language as default
    auto loadOptions = MakeObject<LoadOptions>();
    loadOptions->get_LanguagePreferences()->set_DefaultEditingLanguage(Aspose::Words::EditingLanguage::Russian);

    auto doc = MakeObject<Document>(MyDir + u"No default editing language.docx", loadOptions);

    int localeId = doc->get_Styles()->get_DefaultFont()->get_LocaleId();
    System::Console::WriteLine(localeId == (int)Aspose::Words::EditingLanguage::Russian ? String(u"The document either has no any language set in defaults or it was set to Russian originally.") : String(u"The document default language was set to another than Russian language originally, so it is not overridden."));
    //ExEnd

    ASSERT_EQ((int)Aspose::Words::EditingLanguage::Russian, doc->get_Styles()->get_DefaultFont()->get_LocaleId());

    doc = MakeObject<Document>(MyDir + u"No default editing language.docx");

    ASSERT_EQ((int)Aspose::Words::EditingLanguage::EnglishUS, doc->get_Styles()->get_DefaultFont()->get_LocaleId());
}

namespace gtest_test
{

TEST_F(ExDocument, SetEditingLanguageAsDefault)
{
    s_instance->SetEditingLanguageAsDefault();
}

} // namespace gtest_test

void ExDocument::GetInfoAboutRevisionsInRevisionGroups()
{
    //ExStart
    //ExFor:RevisionGroup
    //ExFor:RevisionGroup.Author
    //ExFor:RevisionGroup.RevisionType
    //ExFor:RevisionGroup.Text
    //ExFor:RevisionGroupCollection
    //ExFor:RevisionGroupCollection.Count
    //ExSummary:Shows how to get info about a group of revisions in document.
    auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

    ASSERT_EQ(7, doc->get_Revisions()->get_Groups()->get_Count());

    // Get info about all of revisions in document
    for (auto group : System::IterateOver(doc->get_Revisions()->get_Groups()))
    {
        System::Console::WriteLine(String::Format(u"Revision author: {0}; Revision type: {1} \n\tRevision text: {2}",group->get_Author(),group->get_RevisionType(),group->get_Text()));
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetInfoAboutRevisionsInRevisionGroups)
{
    s_instance->GetInfoAboutRevisionsInRevisionGroups();
}

} // namespace gtest_test

void ExDocument::GetSpecificRevisionGroup()
{
    //ExStart
    //ExFor:RevisionGroupCollection
    //ExFor:RevisionGroupCollection.Item(Int32)
    //ExSummary:Shows how to get a group of revisions in document.
    auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

    // Get revision group by index
    SharedPtr<RevisionGroup> revisionGroup = doc->get_Revisions()->get_Groups()->idx_get(0);
    //ExEnd

    // Check revision group details
    ASSERT_EQ(Aspose::Words::RevisionType::Deletion, revisionGroup->get_RevisionType());
    ASSERT_EQ(u"Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ", revisionGroup->get_Text());
}

namespace gtest_test
{

TEST_F(ExDocument, GetSpecificRevisionGroup)
{
    s_instance->GetSpecificRevisionGroup();
}

} // namespace gtest_test

void ExDocument::RemovePersonalInformation()
{
    //ExStart
    //ExFor:Document.RemovePersonalInformation
    //ExSummary:Shows how to get or set a flag to remove all user information upon saving the MS Word document.
    auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");
    // If flag sets to 'true' that MS Word will remove all user information from comments, revisions and
    // document properties upon saving the document. In MS Word 2013 and 2016 you can see this using
    // File -> Options -> Trust Center -> Trust Center Settings -> Privacy Options -> then the
    // checkbox "Remove personal information from file properties on save"
    doc->set_RemovePersonalInformation(true);

    // Personal information will not be removed at this time
    // This will happen when we open this document in Microsoft Word and save it manually
    // Once noticeable change will be the revisions losing their author names
    doc->Save(ArtifactsDir + u"Document.RemovePersonalInformation.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.RemovePersonalInformation.docx");
    ASSERT_TRUE(doc->get_RemovePersonalInformation());
}

namespace gtest_test
{

TEST_F(ExDocument, RemovePersonalInformation)
{
    s_instance->RemovePersonalInformation();
}

} // namespace gtest_test

void ExDocument::HideComments()
{
    //ExStart
    //ExFor:LayoutOptions.ShowComments
    //ExSummary:Shows how to show or hide comments in PDF document.
    auto doc = MakeObject<Document>(MyDir + u"Comments.docx");
    doc->get_LayoutOptions()->set_ShowComments(false);

    doc->Save(ArtifactsDir + u"Document.HideComments.pdf");
    //ExEnd

    ASSERT_FALSE(doc->get_LayoutOptions()->get_ShowComments());
}

namespace gtest_test
{

TEST_F(ExDocument, HideComments)
{
    s_instance->HideComments();
}

} // namespace gtest_test

void ExDocument::RevisionOptions()
{
    //ExStart
    //ExFor:ShowInBalloons
    //ExFor:RevisionOptions.ShowInBalloons
    //ExFor:RevisionOptions.CommentColor
    //ExFor:RevisionOptions.DeletedTextColor
    //ExFor:RevisionOptions.DeletedTextEffect
    //ExFor:RevisionOptions.InsertedTextEffect
    //ExFor:RevisionOptions.MovedFromTextColor
    //ExFor:RevisionOptions.MovedFromTextEffect
    //ExFor:RevisionOptions.MovedToTextColor
    //ExFor:RevisionOptions.MovedToTextEffect
    //ExFor:RevisionOptions.RevisedPropertiesColor
    //ExFor:RevisionOptions.RevisedPropertiesEffect
    //ExFor:RevisionOptions.RevisionBarsColor
    //ExFor:RevisionOptions.RevisionBarsWidth
    //ExFor:RevisionOptions.ShowOriginalRevision
    //ExFor:RevisionOptions.ShowRevisionMarks
    //ExFor:RevisionTextEffect
    //ExSummary:Shows how to edit appearance of revisions.
    auto doc = MakeObject<Document>(MyDir + u"Revisions.docx");

    // Get the RevisionOptions object that controls the appearance of revisions
    SharedPtr<Aspose::Words::Layout::RevisionOptions> revisionOptions = doc->get_LayoutOptions()->get_RevisionOptions();

    // Render text inserted while revisions were being tracked in italic green
    revisionOptions->set_InsertedTextColor(Aspose::Words::Layout::RevisionColor::Green);
    revisionOptions->set_InsertedTextEffect(Aspose::Words::Layout::RevisionTextEffect::Italic);

    // Render text deleted while revisions were being tracked in bold red
    revisionOptions->set_DeletedTextColor(Aspose::Words::Layout::RevisionColor::Red);
    revisionOptions->set_DeletedTextEffect(Aspose::Words::Layout::RevisionTextEffect::Bold);

    // In a movement revision, the same text will appear twice: once at the departure point and once at the arrival destination
    // Render the text at the moved-from revision yellow with double strike through and double underlined blue at the moved-to revision
    revisionOptions->set_MovedFromTextColor(Aspose::Words::Layout::RevisionColor::Yellow);
    revisionOptions->set_MovedFromTextEffect(Aspose::Words::Layout::RevisionTextEffect::DoubleStrikeThrough);
    revisionOptions->set_MovedToTextColor(Aspose::Words::Layout::RevisionColor::Blue);
    revisionOptions->set_MovedFromTextEffect(Aspose::Words::Layout::RevisionTextEffect::DoubleUnderline);

    // Render text which had its format changed while revisions were being tracked in bold dark red
    revisionOptions->set_RevisedPropertiesColor(Aspose::Words::Layout::RevisionColor::DarkRed);
    revisionOptions->set_RevisedPropertiesEffect(Aspose::Words::Layout::RevisionTextEffect::Bold);

    // Place a thick dark blue bar on the left side of the page next to lines affected by revisions
    revisionOptions->set_RevisionBarsColor(Aspose::Words::Layout::RevisionColor::DarkBlue);
    revisionOptions->set_RevisionBarsWidth(15.0f);

    // Show revision marks and original text
    revisionOptions->set_ShowOriginalRevision(true);
    revisionOptions->set_ShowRevisionMarks(true);

    // Get movement, deletion, formatting revisions and comments to show up in green balloons on the right side of the page
    revisionOptions->set_ShowInBalloons(Aspose::Words::Layout::ShowInBalloons::Format);
    revisionOptions->set_CommentColor(Aspose::Words::Layout::RevisionColor::BrightGreen);

    // These features are only applicable to formats such as .pdf or .jpg
    doc->Save(ArtifactsDir + u"Document.RevisionOptions.pdf");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, RevisionOptions)
{
    s_instance->RevisionOptions();
}

} // namespace gtest_test

void ExDocument::CopyTemplateStylesViaDocument()
{
    //ExStart
    //ExFor:Document.CopyStylesFromTemplate(Document)
    //ExSummary:Shows how to copies styles from the template to a document via Document.
    auto template_ = MakeObject<Document>(MyDir + u"Rendering.docx");
    auto target = MakeObject<Document>(MyDir + u"Document.docx");

    ASSERT_EQ(18, template_->get_Styles()->get_Count());
    //ExSkip
    ASSERT_EQ(4, target->get_Styles()->get_Count());
    //ExSkip

    target->CopyStylesFromTemplate(template_);
    ASSERT_EQ(18, target->get_Styles()->get_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, CopyTemplateStylesViaDocument)
{
    s_instance->CopyTemplateStylesViaDocument();
}

} // namespace gtest_test

void ExDocument::CopyTemplateStylesViaString()
{
    //ExStart
    //ExFor:Document.CopyStylesFromTemplate(String)
    //ExSummary:Shows how to copies styles from the template to a document via string.
    auto target = MakeObject<Document>(MyDir + u"Document.docx");
    ASSERT_EQ(4, target->get_Styles()->get_Count());
    //ExSkip

    target->CopyStylesFromTemplate(MyDir + u"Rendering.docx");
    ASSERT_EQ(18, target->get_Styles()->get_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, CopyTemplateStylesViaString)
{
    s_instance->CopyTemplateStylesViaString();
}

} // namespace gtest_test

void ExDocument::LayoutCollector()
{
    //ExStart
    //ExFor:Layout.LayoutCollector
    //ExFor:Layout.LayoutCollector.#ctor(Document)
    //ExFor:Layout.LayoutCollector.Clear
    //ExFor:Layout.LayoutCollector.Document
    //ExFor:Layout.LayoutCollector.GetEndPageIndex(Node)
    //ExFor:Layout.LayoutCollector.GetEntity(Node)
    //ExFor:Layout.LayoutCollector.GetNumPagesSpanned(Node)
    //ExFor:Layout.LayoutCollector.GetStartPageIndex(Node)
    //ExFor:Layout.LayoutEnumerator.Current
    //ExSummary:Shows how to see the page spans of nodes.
    // Open a blank document and create a DocumentBuilder
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Create a LayoutCollector object for our document that will have information about the nodes we placed
    auto layoutCollector = MakeObject<Aspose::Words::Layout::LayoutCollector>(doc);

    // The document itself is a node that contains everything, which currently spans 0 pages
    ASPOSE_ASSERT_EQ(doc, layoutCollector->get_Document());
    ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));

    // Populate the document with sections and page breaks
    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    doc->AppendChild(MakeObject<Section>(doc));
    doc->get_LastSection()->AppendChild(MakeObject<Body>(doc));
    builder->MoveToDocumentEnd();
    builder->Write(u"Section 2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    // The collected layout data won't automatically keep up with the real document contents
    ASSERT_EQ(0, layoutCollector->GetNumPagesSpanned(doc));

    // After we clear the layout collection and update it, the layout entity collection will be populated with up-to-date information about our nodes
    // The page span for the document now shows 5, which is what we would expect after placing 4 page breaks
    layoutCollector->Clear();
    doc->UpdatePageLayout();
    ASSERT_EQ(5, layoutCollector->GetNumPagesSpanned(doc));

    // We can also see the start/end pages of any other node, and their overall page spans
    SharedPtr<NodeCollection> nodes = doc->GetChildNodes(Aspose::Words::NodeType::Any, true);
    for (auto node : System::IterateOver(nodes))
    {
        System::Console::WriteLine(String::Format(u"->  NodeType.{0}: ",node->get_NodeType()));
        System::Console::WriteLine(String::Format(u"\tStarts on page {0}, ends on page {1},",layoutCollector->GetStartPageIndex(node),layoutCollector->GetEndPageIndex(node)) + String::Format(u" spanning {0} pages.",layoutCollector->GetNumPagesSpanned(node)));
    }

    // We can iterate over the layout entities using a LayoutEnumerator
    auto layoutEnumerator = MakeObject<Aspose::Words::Layout::LayoutEnumerator>(doc);
    ASSERT_EQ(Aspose::Words::Layout::LayoutEntityType::Page, layoutEnumerator->get_Type());

    // The LayoutEnumerator can traverse the collection of layout entities like a tree
    // We can also point it to any node's corresponding layout entity like this
    layoutEnumerator->set_Current(layoutCollector->GetEntity(doc->GetChild(Aspose::Words::NodeType::Paragraph, 1, true)));
    ASSERT_EQ(Aspose::Words::Layout::LayoutEntityType::Span, layoutEnumerator->get_Type());
    ASSERT_EQ(u"¶", layoutEnumerator->get_Text());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, LayoutCollector)
{
    s_instance->LayoutCollector();
}

} // namespace gtest_test

void ExDocument::LayoutEnumerator()
{
    // Open a document that contains a variety of layout entities
    // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
    // They are defined visually by the rectangular space that they occupy in the document
    auto doc = MakeObject<Document>(MyDir + u"Layout entities.docx");

    // Create an enumerator that can traverse these entities like a tree
    auto layoutEnumerator = MakeObject<Aspose::Words::Layout::LayoutEnumerator>(doc);
    ASPOSE_ASSERT_EQ(doc, layoutEnumerator->get_Document());

    layoutEnumerator->MoveParent(Aspose::Words::Layout::LayoutEntityType::Page);
    ASSERT_EQ(Aspose::Words::Layout::LayoutEntityType::Page, layoutEnumerator->get_Type());

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_20 = [&layoutEnumerator]()
    {
        System::Console::WriteLine(layoutEnumerator->get_Text());
    };

    ASSERT_THROW(_local_func_20(), System::InvalidOperationException);

    // We can call this method to make sure that the enumerator points to the very first entity before we go through it forwards
    layoutEnumerator->Reset();

    // "Visual order" means when moving through an entity's children that are broken across pages,
    // page layout takes precedence and we avoid elements in other pages and move to others on the same page
    System::Console::WriteLine(u"Traversing from first to last, elements between pages separated:");
    TraverseLayoutForward(layoutEnumerator, 1);

    // Our enumerator is conveniently at the end of the collection for us to go through the collection backwards
    System::Console::WriteLine(u"Traversing from last to first, elements between pages separated:");
    TraverseLayoutBackward(layoutEnumerator, 1);

    // "Logical order" means when moving through an entity's children that are broken across pages,
    // node relationships take precedence
    System::Console::WriteLine(u"Traversing from first to last, elements between pages mixed:");
    TraverseLayoutForwardLogical(layoutEnumerator, 1);

    System::Console::WriteLine(u"Traversing from last to first, elements between pages mixed:");
    TraverseLayoutBackwardLogical(layoutEnumerator, 1);
}

namespace gtest_test
{

TEST_F(ExDocument, LayoutEnumerator)
{
    s_instance->LayoutEnumerator();
}

} // namespace gtest_test

void ExDocument::AlwaysCompressMetafiles()
{
    //ExStart
    //ExFor:DocSaveOptions.AlwaysCompressMetafiles
    //ExSummary:Shows how to change metafiles compression in a document while saving.
    // Open a document that contains a Microsoft Equation 3.0 mathematical formula
    auto doc = MakeObject<Document>(MyDir + u"Microsoft equation object.docx");

    // Large metafiles are always compressed when exporting a document in Aspose.Words, but small metafiles are not
    // compressed for performance reason. Some other document editors, such as LibreOffice, cannot read uncompressed
    // metafiles. The following option 'AlwaysCompressMetafiles' was introduced to choose appropriate behavior
    auto saveOptions = MakeObject<DocSaveOptions>();
    // False - small metafiles are not compressed for performance reason
    saveOptions->set_AlwaysCompressMetafiles(false);

    doc->Save(ArtifactsDir + u"Document.AlwaysCompressMetafiles.False.docx", saveOptions);

    // True - all metafiles are compressed regardless of its size
    saveOptions->set_AlwaysCompressMetafiles(true);

    doc->Save(ArtifactsDir + u"Document.AlwaysCompressMetafiles.True.docx", saveOptions);

    ASSERT_TRUE(MakeObject<System::IO::FileInfo>(ArtifactsDir + u"Document.AlwaysCompressMetafiles.True.docx")->get_Length() < MakeObject<System::IO::FileInfo>(ArtifactsDir + u"Document.AlwaysCompressMetafiles.False.docx")->get_Length());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, AlwaysCompressMetafiles)
{
    s_instance->AlwaysCompressMetafiles();
}

} // namespace gtest_test

void ExDocument::CreateNewVbaProject()
{
    //ExStart
    //ExFor:VbaProject.#ctor
    //ExFor:VbaProject.Name
    //ExFor:VbaModule.#ctor
    //ExFor:VbaModule.Name
    //ExFor:VbaModule.Type
    //ExFor:VbaModule.SourceCode
    //ExFor:VbaModuleCollection.Add(VbaModule)
    //ExFor:VbaModuleType
    //ExSummary:Shows how to create a VbaProject from a scratch for using macros.
    auto doc = MakeObject<Document>();

    // Create a new VBA project
    auto project = MakeObject<VbaProject>();
    project->set_Name(u"Aspose.Project");
    doc->set_VbaProject(project);

    // Create a new module and specify a macro source code
    auto module_ = MakeObject<VbaModule>();
    module_->set_Name(u"Aspose.Module");
    // VbaModuleType values:
    // procedural module - A collection of subroutines and functions
    // ------
    // document module - A type of VBA project item that specifies a module for embedded macros and programmatic access
    // operations that are associated with a document
    // ------
    // class module - A module that contains the definition for a new object. Each instance of a class creates
    // a new object, and procedures that are defined in the module become properties and methods of the object
    // ------
    // designer module - A VBA module that extends the methods and properties of an ActiveX control that has been
    // registered with the project
    module_->set_Type(Aspose::Words::VbaModuleType::ProceduralModule);
    module_->set_SourceCode(u"New source code");

    // Add module to the VBA project
    doc->get_VbaProject()->get_Modules()->Add(module_);

    doc->Save(ArtifactsDir + u"Document.CreateVBAMacros.docm");
    //ExEnd

    project = MakeObject<Document>(ArtifactsDir + u"Document.CreateVBAMacros.docm")->get_VbaProject();

    ASSERT_EQ(u"Aspose.Project", project->get_Name());

    SharedPtr<VbaModuleCollection> modules = doc->get_VbaProject()->get_Modules();

    ASSERT_EQ(2, modules->get_Count());

    ASSERT_EQ(u"ThisDocument", modules->idx_get(0)->get_Name());
    ASSERT_EQ(Aspose::Words::VbaModuleType::DocumentModule, modules->idx_get(0)->get_Type());
    ASSERT_TRUE(modules->idx_get(0)->get_SourceCode() == nullptr);

    ASSERT_EQ(u"Aspose.Module", modules->idx_get(1)->get_Name());
    ASSERT_EQ(Aspose::Words::VbaModuleType::ProceduralModule, modules->idx_get(1)->get_Type());
    ASSERT_EQ(u"New source code", modules->idx_get(1)->get_SourceCode());
}

namespace gtest_test
{

TEST_F(ExDocument, CreateNewVbaProject)
{
    s_instance->CreateNewVbaProject();
}

} // namespace gtest_test

void ExDocument::CloneVbaProject()
{
    //ExStart
    //ExFor:VbaProject.Clone
    //ExFor:VbaModule.Clone
    //ExSummary:Shows how to deep clone VbaProject and VbaModule.
    auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");
    auto destDoc = MakeObject<Document>();

    // Clone VbaProject to the document
    SharedPtr<VbaProject> copyVbaProject = doc->get_VbaProject()->Clone();
    destDoc->set_VbaProject(copyVbaProject);

    // In destination document we already have "Module1", because he was cloned with VbaProject
    // Therefore need to remove it before cloning
    SharedPtr<VbaModule> oldVbaModule = destDoc->get_VbaProject()->get_Modules()->idx_get(u"Module1");
    SharedPtr<VbaModule> copyVbaModule = doc->get_VbaProject()->get_Modules()->idx_get(u"Module1")->Clone();
    destDoc->get_VbaProject()->get_Modules()->Remove(oldVbaModule);
    destDoc->get_VbaProject()->get_Modules()->Add(copyVbaModule);

    destDoc->Save(ArtifactsDir + u"Document.CloneVbaProject.docm");
    //ExEnd

    SharedPtr<VbaProject> originalVbaProject = MakeObject<Document>(ArtifactsDir + u"Document.CloneVbaProject.docm")->get_VbaProject();

    ASSERT_EQ(copyVbaProject->get_Name(), originalVbaProject->get_Name());
    ASSERT_EQ(copyVbaProject->get_CodePage(), originalVbaProject->get_CodePage());
    ASPOSE_ASSERT_EQ(copyVbaProject->get_IsSigned(), originalVbaProject->get_IsSigned());
    ASSERT_EQ(copyVbaProject->get_Modules()->get_Count(), originalVbaProject->get_Modules()->get_Count());

    for (int i = 0; i < originalVbaProject->get_Modules()->get_Count(); i++)
    {
        ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_Name(), originalVbaProject->get_Modules()->idx_get(i)->get_Name());
        ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_Type(), originalVbaProject->get_Modules()->idx_get(i)->get_Type());
        ASSERT_EQ(copyVbaProject->get_Modules()->idx_get(i)->get_SourceCode(), originalVbaProject->get_Modules()->idx_get(i)->get_SourceCode());
    }
}

namespace gtest_test
{

TEST_F(ExDocument, CloneVbaProject)
{
    s_instance->CloneVbaProject();
}

} // namespace gtest_test

void ExDocument::ReadMacrosFromExistingDocument()
{
    //ExStart
    //ExFor:Document.VbaProject
    //ExFor:VbaModuleCollection
    //ExFor:VbaModuleCollection.Count
    //ExFor:VbaModuleCollection.Item(System.Int32)
    //ExFor:VbaModuleCollection.Item(System.String)
    //ExFor:VbaModuleCollection.Remove
    //ExFor:VbaModule
    //ExFor:VbaModule.Name
    //ExFor:VbaModule.SourceCode
    //ExFor:VbaProject
    //ExFor:VbaProject.Name
    //ExFor:VbaProject.Modules
    //ExFor:VbaProject.CodePage
    //ExFor:VbaProject.IsSigned
    //ExSummary:Shows how to get access to VBA project information in the document.
    auto doc = MakeObject<Document>(MyDir + u"VBA project.docm");

    // A VBA project inside the document is defined as a collection of VBA modules
    SharedPtr<VbaProject> vbaProject = doc->get_VbaProject();
    ASSERT_TRUE(vbaProject->get_IsSigned());
    //ExSkip
    System::Console::WriteLine(vbaProject->get_IsSigned() ? String::Format(u"Project name: {0} signed; Project code page: {1}; Modules count: {2}\n",vbaProject->get_Name(),vbaProject->get_CodePage(),vbaProject->get_Modules()->LINQ_Count()) : String::Format(u"Project name: {0} not signed; Project code page: {1}; Modules count: {2}\n",vbaProject->get_Name(),vbaProject->get_CodePage(),vbaProject->get_Modules()->LINQ_Count()));

    SharedPtr<VbaModuleCollection> vbaModules = doc->get_VbaProject()->get_Modules();

    ASSERT_EQ(vbaModules->LINQ_Count(), 3);

    for (auto module_ : System::IterateOver(vbaModules))
    {
        System::Console::WriteLine(String::Format(u"Module name: {0};\nModule code:\n{1}\n",module_->get_Name(),module_->get_SourceCode()));
    }

    // Set new source code for VBA module
    // You can retrieve object by integer or by name
    vbaModules->idx_get(0)->set_SourceCode(u"Your VBA code...");
    vbaModules->idx_get(u"Module1")->set_SourceCode(u"Your VBA code...");

    // Remove one of VbaModule from VbaModuleCollection
    vbaModules->Remove(vbaModules->idx_get(2));
    //ExEnd

    ASSERT_EQ(u"AsposeVBAtest", vbaProject->get_Name());
    ASSERT_EQ(2, vbaProject->get_Modules()->LINQ_Count());
    ASSERT_EQ(1251, vbaProject->get_CodePage());
    ASSERT_FALSE(vbaProject->get_IsSigned());

    ASSERT_EQ(u"ThisDocument", vbaModules->idx_get(0)->get_Name());
    ASSERT_EQ(u"Your VBA code...", vbaModules->idx_get(0)->get_SourceCode());

    ASSERT_EQ(u"Module1", vbaModules->idx_get(1)->get_Name());
    ASSERT_EQ(u"Your VBA code...", vbaModules->idx_get(1)->get_SourceCode());
}

namespace gtest_test
{

TEST_F(ExDocument, ReadMacrosFromExistingDocument)
{
    s_instance->ReadMacrosFromExistingDocument();
}

} // namespace gtest_test

void ExDocument::SaveOutputParameters()
{
    //ExStart
    //ExFor:SaveOutputParameters
    //ExFor:SaveOutputParameters.ContentType
    //ExSummary:Shows how to verify Content-Type strings from save output parameters.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // Save the document as a .doc and check parameters
    SharedPtr<Aspose::Words::Saving::SaveOutputParameters> parameters = doc->Save(ArtifactsDir + u"Document.SaveOutputParameters.doc");
    ASSERT_EQ(u"application/msword", parameters->get_ContentType());

    // A .docx or a .pdf will have different parameters
    parameters = doc->Save(ArtifactsDir + u"Document.SaveOutputParameters.pdf");
    ASSERT_EQ(u"application/pdf", parameters->get_ContentType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SaveOutputParameters)
{
    s_instance->SaveOutputParameters();
}

} // namespace gtest_test

void ExDocument::SubDocument()
{
    //ExStart
    //ExFor:SubDocument
    //ExFor:SubDocument.NodeType
    //ExSummary:Shows how to access a master document's subdocument.
    auto doc = MakeObject<Document>(MyDir + u"Master document.docx");

    SharedPtr<NodeCollection> subDocuments = doc->GetChildNodes(Aspose::Words::NodeType::SubDocument, true);
    ASSERT_EQ(1, subDocuments->get_Count());
    //ExSkip

    auto subDocument = System::DynamicCast<Aspose::Words::SubDocument>(subDocuments->idx_get(0));

    // The SubDocument object itself does not contain the documents of the subdocument and only serves as a reference
    ASSERT_FALSE(subDocument->get_IsComposite());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, SubDocument)
{
    s_instance->SubDocument();
}

} // namespace gtest_test

void ExDocument::CreateWebExtension()
{
    //ExStart
    //ExFor:BaseWebExtensionCollection`1.Add(`0)
    //ExFor:TaskPane
    //ExFor:TaskPane.DockState
    //ExFor:TaskPane.IsVisible
    //ExFor:TaskPane.Width
    //ExFor:TaskPane.IsLocked
    //ExFor:TaskPane.WebExtension
    //ExFor:TaskPane.Row
    //ExFor:WebExtension
    //ExFor:WebExtension.Reference
    //ExFor:WebExtension.Properties
    //ExFor:WebExtension.Bindings
    //ExFor:WebExtension.IsFrozen
    //ExFor:WebExtensionReference.Id
    //ExFor:WebExtensionReference.Version
    //ExFor:WebExtensionReference.StoreType
    //ExFor:WebExtensionReference.Store
    //ExFor:WebExtensionPropertyCollection
    //ExFor:WebExtensionBindingCollection
    //ExFor:WebExtensionProperty.#ctor(String, String)
    //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
    //ExFor:WebExtensionStoreType
    //ExFor:WebExtensionBindingType
    //ExFor:TaskPaneDockState
    //ExFor:TaskPaneCollection
    //ExSummary:Shows how to create add-ins inside the document.
    auto doc = MakeObject<Document>();

    // Create taskpane with "MyScript" add-in which will be used by the document
    auto myScriptTaskPane = MakeObject<TaskPane>();
    doc->get_WebExtensionTaskPanes()->Add(myScriptTaskPane);

    // Define task pane location when the document opens
    myScriptTaskPane->set_DockState(Aspose::Words::WebExtensions::TaskPaneDockState::Right);
    myScriptTaskPane->set_IsVisible(true);
    myScriptTaskPane->set_Width(300);
    myScriptTaskPane->set_IsLocked(true);
    // Use this option if you have several task panes
    myScriptTaskPane->set_Row(1);

    // Add "MyScript Math Sample" add-in which will be displayed inside task pane
    SharedPtr<WebExtension> webExtension = myScriptTaskPane->get_WebExtension();

    // Application Id from store
    webExtension->get_Reference()->set_Id(u"WA104380646");
    // The current version of the application used
    webExtension->get_Reference()->set_Version(u"1.0.0.0");
    // Type of marketplace
    webExtension->get_Reference()->set_StoreType(Aspose::Words::WebExtensions::WebExtensionStoreType::OMEX);
    // Marketplace based on your locale
    webExtension->get_Reference()->set_Store(System::Globalization::CultureInfo::get_CurrentCulture()->get_Name());

    webExtension->get_Properties()->Add(MakeObject<WebExtensionProperty>(u"MyScript", u"MyScript Math Sample"));
    webExtension->get_Bindings()->Add(MakeObject<WebExtensionBinding>(u"MyScript", Aspose::Words::WebExtensions::WebExtensionBindingType::Text, u"104380646"));

    // Use this option if you need to block web extension from any action
    webExtension->set_IsFrozen(false);

    doc->Save(ArtifactsDir + u"Document.WebExtension.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Document.WebExtension.docx");
    myScriptTaskPane = doc->get_WebExtensionTaskPanes()->idx_get(0);

    ASSERT_EQ(Aspose::Words::WebExtensions::TaskPaneDockState::Right, myScriptTaskPane->get_DockState());
    ASSERT_TRUE(myScriptTaskPane->get_IsVisible());
    ASPOSE_ASSERT_EQ(300.0, myScriptTaskPane->get_Width());
    ASSERT_TRUE(myScriptTaskPane->get_IsLocked());
    ASSERT_EQ(1, myScriptTaskPane->get_Row());
    webExtension = myScriptTaskPane->get_WebExtension();

    ASSERT_EQ(u"WA104380646", webExtension->get_Reference()->get_Id());
    ASSERT_EQ(u"1.0.0.0", webExtension->get_Reference()->get_Version());
    ASSERT_EQ(Aspose::Words::WebExtensions::WebExtensionStoreType::OMEX, webExtension->get_Reference()->get_StoreType());
    ASSERT_EQ(System::Globalization::CultureInfo::get_CurrentCulture()->get_Name(), webExtension->get_Reference()->get_Store());

    ASSERT_EQ(u"MyScript", webExtension->get_Properties()->idx_get(0)->get_Name());
    ASSERT_EQ(u"MyScript Math Sample", webExtension->get_Properties()->idx_get(0)->get_Value());

    ASSERT_EQ(u"MyScript", webExtension->get_Bindings()->idx_get(0)->get_Id());
    ASSERT_EQ(Aspose::Words::WebExtensions::WebExtensionBindingType::Text, webExtension->get_Bindings()->idx_get(0)->get_BindingType());
    ASSERT_EQ(u"104380646", webExtension->get_Bindings()->idx_get(0)->get_AppRef());

    ASSERT_FALSE(webExtension->get_IsFrozen());
}

namespace gtest_test
{

TEST_F(ExDocument, CreateWebExtension)
{
    s_instance->CreateWebExtension();
}

} // namespace gtest_test

void ExDocument::GetWebExtensionInfo()
{
    //ExStart
    //ExFor:BaseWebExtensionCollection`1
    //ExFor:BaseWebExtensionCollection`1.Add(`0)
    //ExFor:BaseWebExtensionCollection`1.Clear
    //ExFor:BaseWebExtensionCollection`1.GetEnumerator
    //ExFor:BaseWebExtensionCollection`1.Remove(Int32)
    //ExFor:BaseWebExtensionCollection`1.Count
    //ExFor:BaseWebExtensionCollection`1.Item(Int32)
    //ExSummary:Shows how to work with web extension collections.
    auto doc = MakeObject<Document>(MyDir + u"Web extension.docx");
    ASSERT_EQ(1, doc->get_WebExtensionTaskPanes()->get_Count());
    //ExSkip

    // Add new taskpane to the collection
    auto newTaskPane = MakeObject<TaskPane>();
    doc->get_WebExtensionTaskPanes()->Add(newTaskPane);
    ASSERT_EQ(2, doc->get_WebExtensionTaskPanes()->get_Count());
    //ExSkip

    // Enumerate all WebExtensionProperty in a collection
    SharedPtr<WebExtensionPropertyCollection> webExtensionPropertyCollection = doc->get_WebExtensionTaskPanes()->idx_get(0)->get_WebExtension()->get_Properties();
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<WebExtensionProperty>>> enumerator = webExtensionPropertyCollection->GetEnumerator();
        while (enumerator->MoveNext())
        {
            SharedPtr<WebExtensionProperty> webExtensionProperty = enumerator->get_Current();
            System::Console::WriteLine(String::Format(u"Binding name: {0}; Binding value: {1}",webExtensionProperty->get_Name(),webExtensionProperty->get_Value()));
        }
    }

    // We can remove task panes one by one or clear the entire collection
    doc->get_WebExtensionTaskPanes()->Remove(1);
    ASSERT_EQ(1, doc->get_WebExtensionTaskPanes()->get_Count());
    //ExSkip
    doc->get_WebExtensionTaskPanes()->Clear();
    ASSERT_EQ(0, doc->get_WebExtensionTaskPanes()->get_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocument, GetWebExtensionInfo)
{
    s_instance->GetWebExtensionInfo();
}

} // namespace gtest_test

void ExDocument::EpubCover()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");

    // When saving to .epub, some Microsoft Word document properties can be converted to .epub metadata
    doc->get_BuiltInDocumentProperties()->set_Author(u"John Doe");
    doc->get_BuiltInDocumentProperties()->set_Title(u"My Book Title");

    // The thumbnail we specify here can become the cover image
    ArrayPtr<uint8_t> image = System::IO::File::ReadAllBytes(ImageDir + u"Transparent background logo.png");
    doc->get_BuiltInDocumentProperties()->set_Thumbnail(image);

    doc->Save(ArtifactsDir + u"Document.EpubCover.epub");
}

namespace gtest_test
{

TEST_F(ExDocument, EpubCover)
{
    s_instance->EpubCover();
}

} // namespace gtest_test

} // namespace ApiExamples
