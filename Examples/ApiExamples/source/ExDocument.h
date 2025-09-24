#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/text/string_builder.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeChangingArgs.h>
#include <Aspose.Words.Cpp/Model/Nodes/INodeChangingCallback.h>
#include <Aspose.Words.Cpp/Model/Markup/CustomPartCollection.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Markup;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocument : public ApiExampleBase
{
    typedef ExDocument ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Logs the date and time of each node insertion and removal.
    /// Sets a custom font name/size for the text contents of Run nodes.
    /// </summary>
    class HandleNodeChangingFontChanger : public INodeChangingCallback
    {
        typedef HandleNodeChangingFontChanger ThisType;
        typedef INodeChangingCallback BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::String GetLog();
        
        HandleNodeChangingFontChanger();
        
    private:
    
        System::SharedPtr<System::Text::StringBuilder> mLog;
        
        void NodeInserted(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeInserting(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoved(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        void NodeRemoving(System::SharedPtr<Aspose::Words::NodeChangingArgs> args) override;
        
    };
    
    
public:

    void CreateSimpleDocument();
    void Constructor();
    void LoadFromStream();
    void ConvertToPdf();
    void SaveToImageStream();
    void DetectMobiDocumentFormat();
    void OpenFromStreamWithBaseUri();
    void LoadEncrypted();
    void NotSupportedWarning();
    void TempFolder();
    void ConvertToHtml();
    void ConvertToMhtml();
    void ConvertToTxt();
    void ConvertToEpub();
    void SaveToStream();
    void FontChangeViaCallback();
    void AppendDocument();
    void AppendDocumentFromAutomation();
    void ImportList(bool isKeepSourceNumbering);
    void KeepSourceNumberingSameListIds();
    void MergePastedLists();
    void ForceCopyStyles();
    void AdjustSentenceAndWordSpacing();
    void ValidateIndividualDocumentSignatures();
    void DigitalSignature();
    void SignatureValue();
    void AppendAllDocumentsInFolder();
    void JoinRunsWithSameFormatting();
    void DefaultTabStop();
    void CloneDocument();
    void DocumentGetTextToString();
    void ProtectUnprotect();
    void DocumentEnsureMinimum();
    void RemoveMacrosFromDocument();
    void GetPageCount();
    void GetUpdatedPageProperties();
    void GetOriginalFileInfo();
    void FootnoteColumns();
    void RemoveExternalSchemaReferences();
    void UpdateThumbnail();
    void HyphenationOptions();
    void HyphenationOptionsDefaultValues();
    void HyphenationZoneException();
    void OoxmlComplianceVersion();
    void ImageSaveOptions();
    void Cleanup();
    void AutomaticallyUpdateStyles();
    void DefaultTemplate();
    void UseSubstitutions();
    void SetInvalidateFieldTypes();
    void LayoutOptionsHiddenText(bool showHiddenText);
    void LayoutOptionsParagraphMarks(bool showParagraphMarks);
    void UpdatePageLayout();
    void DocPackageCustomParts();
    void ShadeFormData(bool useGreyShading);
    void VersionsCount();
    void WriteProtection();
    void RemovePersonalInformation(bool saveWithoutPersonalInfo);
    void ShowComments();
    void CopyTemplateStylesViaDocument();
    void CopyTemplateStylesViaDocumentNew();
    void ReadMacrosFromExistingDocument();
    void SaveOutputParameters();
    void SubDocument();
    void CreateWebExtension();
    void GetWebExtensionInfo();
    void EpubCover();
    void TextWatermark();
    void ImageWatermark();
    void ImageWatermarkStream();
    void SpellingAndGrammarErrors(bool showErrors);
    void IgnorePrinterMetrics();
    void ExtractPages();
    void SpellingOrGrammar(bool checkSpellingGrammar);
    void AllowEmbeddingPostScriptFonts();
    void Frameset();
    void OpenAzw();
    void OpenEpub();
    void OpenXml();
    void MoveToStructuredDocumentTag();
    void IncludeTextboxesFootnotesEndnotesInStat();
    void SetJustificationMode();
    void PageIsInColor();
    void InsertDocumentInline();
    void SaveDocumentToStream(Aspose::Words::SaveFormat saveFormat);
    void HasMacros();
    void PunctuationKerning();
    void RemoveBlankPages();
    void ExtractPagesWithOptions();
    
protected:

    static void TestFontChangeViaCallback(System::String log);
    static void TestDocPackageCustomParts(System::SharedPtr<Aspose::Words::Markup::CustomPartCollection> parts);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


