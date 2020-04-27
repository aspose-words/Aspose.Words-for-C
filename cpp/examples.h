#pragma once

#include "RunExamples.h"

// =====================================================
// =====================================================
// Loading and Saving
// =====================================================
// =====================================================
void AccessAndVerifySignature();
void CheckFormat();
void ConvertDocumentToByte();
void ConvertDocumentToEPUB();
void ConvertDocumentToHtmlWithRoundtrip();
void ConvertDocumentToPCL();
void CreateDocument();
void DetectDocumentSignatures();
void DigitallySignedPdf();
void DigitallySignedPdfUsingCertificateHolder();
void Doc2Pdf();
void ExportFontsAsBase64();
void ExportResourcesUsingHtmlSaveOptions();
//void ImageToPdf(); // System::Drawing::Image::get_FrameDimensionsList() isn't implemented in the asposecpp lib
void Load_Options();
void LoadAndSaveHtmlFormFieldAsContentControlInDOCX();
void LoadAndSaveToDisk();
void LoadAndSaveToStream();
void LoadTxt();
void OpenEncryptedDocument();
void PageSplitter();
void SaveDocWithHtmlSaveOptions();
void SaveOptionsHtmlFixed();
void SpecifySaveOption();
void SplitIntoHtmlPages();
void WorkingWithDoc(); // Source document is missing
void WorkingWithOoxml();
void WorkingWithRTF();
void WorkingWithTxt(); // Source document is missing
void WorkingWithVbaMacros();

// =====================================================
// =====================================================
// Mail-Merge
// =====================================================
// =====================================================
//void ApplyCustomLogicToEmptyRegions(); // isn't compiled due to using of DataSet
void ExecuteArray();
//void ExecuteWithRegionsDataTable(); // isn't compiled due to using of OleDbXXX & DataTable
void HandleMailMergeSwitches();
//void LINQtoXMLMailMerge(); // isn't compiled due to using of Linq2Xml
//void MailMergeAlternatingRows(); // isn't compiled due to using of DataTable
void MailMergeAndConditionalField();
void MailMergeCleanUp();
void MailMergeFormFields();
//void MailMergeImageFromBlob(); // isn't compiled due to using of OleDbXXX & IDataReader
//void MailMergeUsingMustacheSyntax.Run(); // isn't compiled due to using of DataSet
//void MultipleDocsInMailMerge(); // isn't compiled due to using of OleDbXXX & DataTable
//void NestedMailMerge(); // isn't compiled due to using of DataSet
void NestedMailMergeCustom();
//void ProduceMultipleDocuments(); // isn't compiled due to using of OleDbXXX & DataTable
//void RemoveEmptyRegions(); // isn't compiled due to using of DataSet
//void RemoveRowsFromTable(); // isn't compiled due to using of DataSet
void SimpleMailMerge();
//void XMLMailMerge(); // isn't compiled due to using of DataSet
void MailMergeImageField();

// =====================================================
// =====================================================
// Programming with Documents
// =====================================================
// =====================================================

// Bookmarks
// =====================================================
void AccessBookmarks();
void BookmarkNameAndText();
void BookmarkTable();
void CopyBookmarkedText();
void CreateBookmark();
void UntangleRowBookmarks();
void ShowHideBookmarks();

// Charts
// =====================================================
void ChartNumberFormat();
void CreateChartUsingShape();
void CreateColumnChart();
void InsertAreaChart();
void InsertBubbleChart();
void InsertScatterChart();
void WorkingWithChartAxis();
void WorkWithChartDataLabels();
void WorkWithSingleChartDataPoint();
void WorkWithSingleChartSeries();

// Comments
// =====================================================
void AddComments();
void AnchorComment();
void CommentReply();
void ProcessComments();
void RemoveRegionText();

// ConvertUtil
// =====================================================
void UtilityClasses();

// Document
// =====================================================
void AccessStyles();
void AddGroupShapeToDocument();
void CheckBoxTypeContentControl();
void CheckDMLTextEffect();
void CleansUnusedStylesandLists();
void CloningDocument();
void ComboBoxContentControl();
void CompareDocument();
void CreateHeaderFooterUsingDocBuilder();
void DocProperties(); // exception at execution
void DocumentBuilderBuildTable();
void DocumentBuilderHorizontalRule();
void DocumentBuilderInsertBookmark();
void DocumentBuilderInsertBreak();
void DocumentBuilderInsertElements();
void DocumentBuilderInsertImage();
void DocumentBuilderInsertParagraph();
void DocumentBuilderInsertTCField();
void DocumentBuilderInsertTCFieldsAtText();
void DocumentBuilderInsertTOC();
void DocumentBuilderMovingCursor();
void DocumentBuilderSetFormatting();
void DocumentPageSetup();
void ExtractContentBetweenBlockLevelNodes();
void ExtractContentBetweenBookmark();
void ExtractContentBetweenCommentRange();
void ExtractContentBetweenParagraphs();
void ExtractContentBetweenParagraphStyles();
void ExtractContentBetweenRuns();
void ExtractContentUsingDocumentVisitor();
void ExtractContentUsingField();
void ExtractTableOfContents();
void ExtractTextOnly();
//void GenerateACustomBarCodeImage.Run(); // Depend from Aspose.BarCode
void GetFontLineSpacing();
void GetVariables();
void InsertDoc();
void PageNumbersOfNodes();
void ParagraphStyleSeparator();
void ProtectDocument();
void RemoveBreaks();
void RemoveFooters();
void RemoveTOCFromDocument();
void RichTextBoxContentControl();
void SetCompatibilityOptions();
void Setuplanguagepreferences();
void SetViewOption();
void SigningSignatureLine();
void UpdateContentControls();
void WriteAndFont();
//void WorkingWithImportFormatOptions(); // Source document is missing
void WorkingWithMarkdownFeatures();
void WorkingWithRevisions();
void WorkingWithRtfSaveOptions();
void WorkingWithSaveOptions();
void WorkingWithWebExtension();

// EndNote and Footnote
// =====================================================
void WorkingWithFootnote();

// Fields
// =====================================================
void ChangeFieldUpdateCultureSource();
void ChangeLocale();
void ConvertFieldsInBody();
void ConvertFieldsInDocument();
void ConvertFieldsInParagraph();
void EvaluateIFCondition();
void FieldUpdateCulture();
void FormatFieldResult();
void FormFieldsGetByName();
void FormFieldsGetFormFieldsCollection();
void FormFieldsWorkWithProperties();
void GetFieldNames();
void InsertAdvanceFieldWithoutDocumentBuilder();
void InsertASKFieldWithoutDocumentBuilder();
void InsertAuthorField();
void InsertFields();
void InsertFieldNone();
void InsertFormFields();
void InsertIncludeFieldWithoutDocumentBuilder();
void InsertMailMergeAddressBlockFieldUsingDOM();
void InsertMergeFieldUsingDOM();
void InsertNestedFields();
void InsertTOAFieldWithoutDocumentBuilder();
void RemoveField();
void RenameMergeFields();
void SpecifylocaleAtFieldlevel();
void UpdateDocFields();
void UseOfficeMathProperties();

// Find and Replace
// =====================================================
void FindAndHighlight();
void FindReplaceUsingMetaCharacters();
void ReplaceHtmlTextWithMetaCharacters();
void ReplaceTextWithField();
void ReplaceWithEvaluator();
void ReplaceWithHTML();
void ReplaceWithRegex();
void ReplaceWithString();
void UsingLegacyOrder();
void IgnoreText();

// Hyperlink
// =====================================================
void ReplaceHyperlinks();

// Images
// =====================================================
void AddImageToEachPage();
void AddWatermark();
void CompressImages();
void ExtractImagesToFiles();
void InsertBarcodeImage();
void RemoveWatermark();

// Joining and Appending
// =====================================================
void AppendDocumentManually();
//void AppendWithImportFormatOptions(); // Source document is missing
void BaseDocument();
void ConvertNumPageFields();
void DifferentPageSetup();
void JoinContinuous();
void JoinNewPage();
void KeepSourceFormatting();
void KeepSourceTogether();
void LinkHeadersFooters();
void ListKeepSourceFormatting();
void ListUseDestinationStyles();
void PrependDocument();
void RemoveSourceHeadersFooters();
void RestartPageNumbering();
void SimpleAppendDocument();
void UnlinkHeadersFooters();
void UpdatePageLayout();
void UseDestinationStyles();

// Linked TextBoxes
// =====================================================
void WorkingWithLinkedTextboxes();

// List
// =====================================================
void WorkingWithList();

// Node
// =====================================================
void ExNode();

// Ranges
// =====================================================
void RangesDeleteText();
void RangesGetText();

// Sections
// =====================================================
void AddDeleteSection();
void AppendSectionContent();
void CloneSection();
void CopySection();
void DeleteHeaderFooterContent();
void DeleteSectionContent();
//void ModifyPageSetupInAllSectionsOfDocument(); // Source document is missing
void SectionsAccessByIndex();

// Shapes
// =====================================================
void WorkingWithShapes(); // Source document is missing

// StructuredDocumentTag 
// =====================================================
void WorkingWithSDT();

// Styles
// =====================================================
void ChangeStyleOfTOCLevel();
void ChangeTOCTabStops();
void CopyStyles();
void ExtractContentBasedOnStyles();
void InsertStyleSeparator();

// Tables
// =====================================================
void AddRemoveColumn();
void ApplyFormatting();
void ApplyStyle();
void AutoFitTableToContents();
void AutoFitTableToFixedColumnWidths();
void AutoFitTableToWindow();
//void BuildTableFromDataTable(); // isn't compiled due to using of DataTable
void CloneTable();
void ExtractOrReplaceText();
void FindingIndex();
void InsertTableDirectly();
void InsertTableFromHtml();
void InsertTableUsingDocumentBuilder();
void JoiningAndSplittingTable();
void KeepTablesAndRowsBreaking();
void MergedCells(); // raised NullReferenceException
void RepeatRowsOnSubsequentPages();
void SpecifyHeightAndWidth();
void TablePosition();

// Theme
// =====================================================
void ManipulateThemeProperties();

// =====================================================
// =====================================================
// Quick Start
// =====================================================
// =====================================================
void AppendDocuments();
void ApplyLicense();
void ApplyLicenseFromFile();
void ApplyLicenseFromStream();
//void ApplyMeteredLicense(); // Metered license isn't supported
void FindAndReplace();
void HelloWorld();
void UpdateFields();
void WorkingWithNodes();

// =====================================================
// =====================================================
// Rendering and Printing
// =====================================================
// =====================================================
void DocumentLayoutHelper();
void EmbeddedFontsInPDF();
void EmbeddingWindowsStandardFonts();
void EnumerateLayoutElements();
void HyphenateWordsOfLanguages();
void ImageColorFilters(); // saving into images don't work properly
void LoadHyphenationDictionaryForLanguage();
//void PrintProgressDialog(); // using of GUI
//void Print_CachePrinterSettings; // using of GUI
//void ReadActiveXControlProperties(); // doesn't work due to using of OleControl
void ReceiveNotificationsOfFont();
void RenderShape();
void ResourceSteamFontSource();
//void SaveAsMultipageTiff(); // saving into images don't work properly
void WorkingWithFontSettings();
void SetFontsFoldersMultipleFolders();
void SetFontsFoldersSystemAndCustomFolder();
void SetHorizontalAndVerticalImageResolution();
void SetTrueTypeFontsFolder();
void SpecifyDefaultFontWhenRendering();
//WorkingWithFontResolution(); // Source document is missing
void WorkingWithFontSources();
void WorkingWithPdfSaveOptions();
void SetFontsFolders();
void SetFontsFoldersDefaultInstance();
void SetFontsFoldersWithPriority();
