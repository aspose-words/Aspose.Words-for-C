#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <system/collections/list.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Saving/IFontSavingCallback.h>
#include <Aspose.Words.Cpp/Model/Nodes/INodeChangingCallback.h>
#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/FindReplace/IReplacingCallback.h>
#include <Aspose.Words.Cpp/Model/Document/IWarningCallback.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Loading { enum class ResourceLoadingAction; } } }
namespace Aspose { namespace Words { namespace Loading { class ResourceLoadingArgs; } } }
namespace Aspose { namespace Words { class WarningInfo; } }
namespace Aspose { namespace Words { namespace Saving { class FontSavingArgs; } } }
namespace Aspose { namespace Words { class NodeChangingArgs; } }
namespace Aspose { namespace Words { class Revision; } }
namespace Aspose { namespace Words { enum class NodeType; } }
namespace Aspose { namespace Words { namespace Replacing { enum class ReplaceAction; } } }
namespace Aspose { namespace Words { namespace Replacing { class ReplacingArgs; } } }
namespace Aspose { namespace Words { class Document; } }
namespace Aspose { namespace Words { namespace Markup { class CustomPartCollection; } } }
namespace Aspose { namespace Words { namespace Layout { class LayoutEnumerator; } } }

namespace ApiExamples {

class ExDocument : public ApiExampleBase
{
public:

    /// <summary>
    /// Prints information about fonts and saves them alongside their output .html.
    /// </summary>
    class HandleFontSaving : public Aspose::Words::Saving::IFontSavingCallback
    {
        typedef HandleFontSaving ThisType;
        typedef Aspose::Words::Saving::IFontSavingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void FontSaving(System::SharedPtr<Aspose::Words::Saving::FontSavingArgs> args) override;
        
    };
    
    class HandleNodeChangingFontChanger : public Aspose::Words::INodeChangingCallback
    {
        typedef HandleNodeChangingFontChanger ThisType;
        typedef Aspose::Words::INodeChangingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void NodeInserted(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeInserting(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoved(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoving(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        
    };
    
    
private:

    /// <summary>
    /// Resource loading callback that, upon encountering external resources,
    /// acknowledges CSS style sheets and replaces all images with a substitute.
    /// </summary>
    class HtmlLinkedResourceLoadingCallback : public Aspose::Words::Loading::IResourceLoadingCallback
    {
        typedef HtmlLinkedResourceLoadingCallback ThisType;
        typedef Aspose::Words::Loading::IResourceLoadingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        Aspose::Words::Loading::ResourceLoadingAction ResourceLoading(System::SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args) override;
        
    };
    
    /// <summary>
    /// IWarningCallback that prints warnings and their details as they arise during document loading.
    /// </summary>
    class DocumentLoadingWarningCallback : public Aspose::Words::IWarningCallback
    {
        typedef DocumentLoadingWarningCallback ThisType;
        typedef Aspose::Words::IWarningCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        void Warning(System::SharedPtr<Aspose::Words::WarningInfo> info) override;
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> GetWarnings();
        
        DocumentLoadingWarningCallback();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    private:
    
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> mWarnings;
        
    };
    
    class UseLegacyOrderReplacingCallback : public Aspose::Words::Replacing::IReplacingCallback
    {
        typedef UseLegacyOrderReplacingCallback ThisType;
        typedef Aspose::Words::Replacing::IReplacingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::SharedPtr<System::Collections::Generic::List<System::String>> Matches;
        
        Aspose::Words::Replacing::ReplaceAction Replacing(System::SharedPtr<Aspose::Words::Replacing::ReplacingArgs> e) override;
        
        UseLegacyOrderReplacingCallback();
        
    protected:
    
        System::Object::shared_members_type GetSharedMembers() override;
        
    };
    
    
public:

    void LicenseFromFileNoPath();
    void LicenseFromStream();
    void LoadOptionsCallback();
    void DocumentCtor();
    void ConvertToPdf();
    void OpenAndSaveToFile();
    void OpenFromStream();
    void OpenFromStreamWithBaseUri();
    void OpenDocumentFromWeb();
    void InsertHtmlFromWebPage();
    void LoadFormat();
    void LoadEncrypted();
    void ConvertShapeToOfficeMath(bool isConvertShapeToOfficeMath);
    void LoadOptionsEncoding();
    void LoadOptionsFontSettings();
    void LoadOptionsMswVersion();
    void LoadOptionsWarningCallback();
    void ConvertToHtml();
    void ConvertToMhtml();
    void ConvertToTxt();
    void SaveToStream();
    void Doc2EpubSave();
    void Doc2EpubSaveOptions();
    void DownsampleOptions();
    void SaveHtmlPrettyFormat(bool isPrettyFormat);
    void SaveHtmlWithOptions();
    void SaveHtmlExportFonts();
    void FontChangeViaCallback();
    void AppendDocument();
    void AppendDocumentFromAutomation();
    void ValidateIndividualDocumentSignatures();
    void DigitalSignature();
    void AppendAllDocumentsInFolder();
    void JoinRunsWithSameFormatting();
    void DefaultTabStop();
    void CloneDocument();
    void ChangeFieldUpdateCultureSource();
    void DocumentGetTextToString();
    void DocumentByteArray();
    void Protect();
    void DocumentEnsureMinimum();
    void RemoveMacrosFromDocument();
    void UpdateTableLayout();
    void GetPageCount();
    void GetUpdatedPageProperties();
    void TableStyleToDirectFormatting();
    void GetOriginalFileInfo();
    void FootnoteColumns();
    void Footnotes();
    void Endnotes();
    void Compare();
    void CompareDocumentWithRevisions();
    void CompareOptions();
    void RemoveExternalSchemaReferences();
    void RemoveUnusedResources();
    void StartTrackRevisions();
    void ShowRevisionBalloons();
    void AcceptAllRevisions();
    void GetRevisedPropertiesOfList();
    void UpdateThumbnail();
    void HyphenationOptions();
    void HyphenationOptionsDefaultValues();
    void HyphenationOptionsExceptions();
    void ExtractPlainTextFromDocument();
    void GetPlainTextBuiltInDocumentProperties();
    void GetPlainTextCustomDocumentProperties();
    void ExtractPlainTextFromStream();
    void OoxmlComplianceVersion();
    void ImageSaveOptions();
    void CleanUpStyles();
    void Revisions();
    void RevisionCollection();
    void AutomaticallyUpdateStyles();
    void DefaultTemplate();
    void Sections();
    void UseLegacyOrder(bool isUseLegacyOrder);
    void SetInvalidateFieldTypes();
    void LayoutOptions();
    void MailMergeSettings();
    void OdsoEmail();
    void MailingLabelMerge();
    void OdsoFieldMapDataCollection();
    void OdsoRecipientDataCollection();
    void DocPackageCustomParts();
    void ShadeFormData();
    void VersionsCount();
    void WriteProtection();
    void AddEditingLanguage();
    void SetEditingLanguageAsDefault();
    void GetInfoAboutRevisionsInRevisionGroups();
    void GetSpecificRevisionGroup();
    void RemovePersonalInformation();
    void HideComments();
    void RevisionOptions();
    void CopyTemplateStylesViaDocument();
    void CopyTemplateStylesViaString();
    void LayoutCollector();
    void LayoutEnumerator();
    void AlwaysCompressMetafiles();
    void CreateNewVbaProject();
    void CloneVbaProject();
    void ReadMacrosFromExistingDocument();
    void SaveOutputParameters();
    void SubDocument();
    void CreateWebExtension();
    void GetWebExtensionInfo();
    void EpubCover();
    
protected:

    static void TestLoadOptionsWarningCallback(System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::WarningInfo>>> warnings);
    /// <summary>
    /// Returns true if the passed revision has a parent node with the type specified by parentType
    /// </summary>
    bool HasParentOfType(System::SharedPtr<Aspose::Words::Revision> revision, Aspose::Words::NodeType parentType);
    static void CheckUseLegacyOrderResults(bool isUseLegacyOrder, System::SharedPtr<ExDocument::UseLegacyOrderReplacingCallback> callback);
    void TestOdsoEmail(System::SharedPtr<Aspose::Words::Document> doc);
    static void TestDocPackageCustomParts(System::SharedPtr<Aspose::Words::Markup::CustomPartCollection> parts);
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order.
    /// </summary>
    static void TraverseLayoutForward(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order.
    /// </summary>
    static void TraverseLayoutBackward(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order.
    /// </summary>
    static void TraverseLayoutForwardLogical(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order.
    /// </summary>
    static void TraverseLayoutBackwardLogical(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t depth);
    /// <summary>
    /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent.
    /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document.
    /// </summary>
    static void PrintCurrentEntity(System::SharedPtr<Aspose::Words::Layout::LayoutEnumerator> layoutEnumerator, int32_t indent);
    
};

} // namespace ApiExamples


