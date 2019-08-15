#pragma once

#include "RunExamples.h"

// Quick Start
void AppendDocuments();
void ApplyLicense();
void FindAndReplace();
void HelloWorld();
void WorkingWithNodes();
void UpdateFields();
void ApplyLicenseFromFile();
void ApplyLicenseFromStream();

// Loading and Saving
void LoadAndSaveToDisk();
void LoadAndSaveToStream();
void CreateDocument();
void CheckFormat();
void ConvertDocumentToByte();
void LoadTxt();
void DetectDocumentSignatures();
void WorkingWithTxt();
void WorkingWithRTF();
void Load_Options();
void PageSplitter();
void ConvertDocumentToEPUB();
void ConvertDocumentToHtmlWithRoundtrip();
void Doc2Pdf();
void ExportFontsAsBase64();
void ExportResourcesUsingHtmlSaveOptions();
void LoadAndSaveHtmlFormFieldAsContentControlInDOCX();
void SaveDocWithHtmlSaveOptions();
void SpecifySaveOption();
void SplitIntoHtmlPages();
void WorkingWithOoxml();

// =================================
// Programming with Documents
// =================================

// Joining and Appending
void SimpleAppendDocument();
void KeepSourceFormatting();
void UseDestinationStyles();
void JoinContinuous();
void JoinNewPage();
void RestartPageNumbering();
void LinkHeadersFooters();
void UnlinkHeadersFooters();
void RemoveSourceHeadersFooters();
void DifferentPageSetup();
void ListUseDestinationStyles();
void ListKeepSourceFormatting();
void KeepSourceTogether();
void BaseDocument();
void AppendDocumentManually();
void PrependDocument();
void ConvertNumPageFields();
void UpdatePageLayout();

// Find and Replace
void FindAndHighlight();
void ReplaceWithString();
void ReplaceWithRegex();
void ReplaceWithEvaluator();
void FindReplaceUsingMetaCharacters();
void ReplaceTextWithField();
void ReplaceHtmlTextWithMetaCharacters();

// Bookmarks
void CopyBookmarkedText();
void UntangleRowBookmarks();
void BookmarkTable();
void BookmarkNameAndText();
void AccessBookmarks();
void CreateBookmark();

// Shapes
void WorkingWithShapes();

// Comments
void ProcessComments();
void AddComments();
void AnchorComment();
void CommentReply();
void RemoveRegionText();

// ConvertUtil
void UtilityClasses();

// Document
void ExtractContentBetweenParagraphs();
void ExtractContentBetweenBlockLevelNodes();
void ExtractContentBetweenParagraphStyles();
void ExtractContentBetweenRuns();
void ExtractContentBetweenBookmark();
void ExtractContentBetweenCommentRange();
void RemoveBreaks();
void CloningDocument();
void CompareDocument();
void ProtectDocument();
void AccessStyles();
void GetVariables();
void SetViewOption();
void CreateHeaderFooterUsingDocBuilder();
void RemoveFooters();
void AddGroupShapeToDocument();
void DocProperties();
void WriteAndFont();
void DocumentBuilderInsertParagraph();
void DocumentBuilderBuildTable();
void DocumentBuilderInsertBreak();
void DocumentBuilderInsertImage();
void DocumentBuilderInsertBookmark();
void DocumentBuilderInsertElements();
void DocumentBuilderSetFormatting();
void DocumentBuilderMovingCursor();
void InsertDoc();
void DocumentBuilderInsertTCFieldsAtText();
void RemoveTOCFromDocument();
void SetCompatibilityOptions();
void Setuplanguagepreferences();
void WorkingWithRevisions();
void ExtractContentUsingField();
void ExtractTextOnly();
void DocumentBuilderInsertTOC();
void DocumentBuilderInsertTCField();
void CheckBoxTypeContentControl();
void RichTextBoxContentControl();
void ComboBoxContentControl();
void UpdateContentControls();
void CleansUnusedStylesandLists();
void DocumentPageSetup();
void WorkingWithSaveOptions();
void DocumentBuilderInsertHorizontalRule();
void PageNumbersOfNodes();

// Fields
void RemoveField();
void ConvertFieldsInDocument();
void ConvertFieldsInBody();
void ConvertFieldsInParagraph();
void SpecifylocaleAtFieldlevel();
void InsertFormFields();
void FormFieldsGetFormFieldsCollection();
void FormFieldsGetByName();
void FormFieldsWorkWithProperties();
void RenameMergeFields();
void InsertNestedFields();
void ChangeLocale();
void UpdateDocFields();
void UseOfficeMathProperties();
void InsertField();
void InsertMergeFieldUsingDOM();
void InsertMailMergeAddressBlockFieldUsingDOM();
void InsertAdvanceFieldWithoutDocumentBuilder();
void InsertASKFieldWithoutDocumentBuilder();
void InsertAuthorField();
void GetFieldNames();
void InsertTOAFieldWithoutDocumentBuilder();
void InsertIncludeTextFieldWithoutDocumentBuilder();
void EvaluateIFCondition();
void FieldUpdateCulture();
void FormatFieldResult();
void InsertFieldNone();

// Images
void AddWatermark();
void RemoveWatermark();
void CompressImages();
void ExtractImagesToFiles();
void InsertBarcodeImage();

// Ranges
void RangesGetText();
void RangesDeleteText();

// Charts
void ChartNumberFormat();
void CreateChartUsingShape();
void CreateColumnChart();
void InsertAreaChart();
void InsertBubbleChart();
void InsertScatterChart();
void WorkWithChartDataLabel();
void WorkWithSingleChartDataPoint();
void WorkWithSingleChartSeries();
void WorkingWithChartAxis();

// Theme
void ManipulateThemeProperties();

// Node
void ExNode();

// Hyperlink
void ReplaceHyperlinks();

// Styles
void ExtractContentBasedOnStyles();
void ChangeStyleOfTOCLevel();
void ChangeTOCTabStops();
void InsertStyleSeparator();
void CopyStyles();

// Tables
void AddRemoveColumn();
void ApplyFormatting();
void ApplyStyle();
void AutoFitTableToContents();
void AutoFitTableToFixedColumnWidths();
void AutoFitTableToWindow();
void CloneTable();
void ExtractOrReplaceText();
void FindingIndex();
void InsertTableDirectly();
void InsertTableUsingDocumentBuilder();
void JoiningAndSplittingTable();
void KeepTablesAndRowsBreaking();
void MergedCells();
void RepeatRowsOnSubsequentPages();
void SpecifyHeightAndWidth();
void TablePosition();
void InsertTableFromHtml();

// Sections
void AddDeleteSection();
void AppendSectionContent();
void CloneSection();
void CopySection();
void DeleteHeaderFooterContent();
void DeleteSectionContent();
void SectionsAccessByIndex();

// StructuredDocumentTag 
void WorkingWithSDT();

// EndNote and Footnote
void WorkingWithFootnote();

// List
void WorkingWithList();

// =====================================================
// Mail-Merge
// =====================================================
void SimpleMailMerge();
void MailMergeFormFields();
void ExecuteArray();
void NestedMailMergeCustom();
void HandleMailMergeSwitches();
void MailMergeAndConditionalField();
void MailMergeCleanUp();

// =====================================================
// Rendering and Printing
// =====================================================
void DocumentLayoutHelper();
void EmbeddedFontsInPDF();
void EmbeddingWindowsStandardFonts();
void EnumerateLayoutElements();
void HyphenateWordsOfLanguages();
void LoadHyphenationDictionaryForLanguage();
void ReceiveNotificationsOfFont();
void RenderShape();
void SetFontSettings();
void SetFontsFoldersMultipleFolders();
void SetFontsFoldersSystemAndCustomFolder();
void SetHorizontalAndVerticalImageResolution();
void SetTrueTypeFontsFolder();
void SpecifyDefaultFontWhenRendering();
void WorkingWithFontSources();
void WorkingWithPdfSaveOptions();