#pragma once

#include "RunExamples.h"

// Quick Start
void ApplyLicense();
void ApplyLicenseFromFile();
void ApplyLicenseFromStream();
// TODO(std_string) : ApplyMeteredLicense isn't compiled
//void ApplyMeteredLicense();
void AppendDocuments();
void FindAndReplace();
void HelloWorld();
void UpdateFields();
void WorkingWithNodes();

// Loading and Saving
// TODO (std_string) : signatures isn't supported
//void AccessAndVerifySignature();
void CheckFormat();
void ConvertDocumentToByte();
void ConvertDocumentToEPUB();
void ConvertDocumentToHtmlWithRoundtrip();
void ConvertDocumentToPCL();
void CreateDocument();
void DetectDocumentSignatures();
// TODO (std_string) : doesn't work in C# code (System.NullReferenceException raised)
void DigitallySignedPdf();
// TODO (std_string) : signatures isn't supported
//void DigitallySignedPdfUsingCertificateHolder();
void Doc2Pdf();
void ExportFontsAsBase64();
void ExportResourcesUsingHtmlSaveOptions();
// TODO (std_string) : System::Drawing::Image::get_FrameDimensionsList() isn't implemented in the asposecpp lib
//void ImageToPdf();
void LoadAndSaveHtmlFormFieldAsContentControlInDOCX();
void LoadAndSaveToDisk();
void LoadAndSaveToStream();
void Load_Options();
void LoadTxt();
// TODO (std_string) : doesn't work properly, but raises FileCorruptedException instead of NotSupportedException / NotImplemented Exception
//void OpenEncryptedDocument();
void PageSplitter();
void SaveDocWithHtmlSaveOptions();
void SaveOptionsHtmlFixed();
void SpecifySaveOption();
void SplitIntoHtmlPages();
// TODO(std_string) : encryption doesn't work
void WorkingWithDoc();
void WorkingWithOoxml();
void WorkingWithRTF();
void WorkingWithTxt();
void WorkingWithVbaMacros();

// =================================
// Programming with Documents
// =================================

// Joining and Appending
void AppendDocumentManually();
// TODO(std_string) : absent documents
//void AppendWithImportFormatOptions();
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

// Find and Replace
void FindAndHighlight();
void FindReplaceUsingMetaCharacters();
void ReplaceHtmlTextWithMetaCharacters();
void ReplaceTextWithField();
void ReplaceWithEvaluator();
void ReplaceWithHTML();
void ReplaceWithRegex();
void ReplaceWithString();

// Bookmarks
void AccessBookmarks();
void BookmarkNameAndText();
void BookmarkTable();
void CopyBookmarkedText();
void CreateBookmark();
void UntangleRowBookmarks();

// Shapes
void WorkingWithShapes();

// Comments
void AddComments();
void AnchorComment();
void CommentReply();
void ProcessComments();
void RemoveRegionText();

// ConvertUtil
void UtilityClasses();

// Document
void AccessStyles();
void AddGroupShapeToDocument();
void CheckBoxTypeContentControl();
void CheckDMLTextEffect();
void CleansUnusedStylesandLists();
void CloningDocument();
void ComboBoxContentControl();
void CompareDocument();
void CreateHeaderFooterUsingDocBuilder();
void DocProperties();
void DocumentBuilderBuildTable();
void DocumentBuilderInsertBookmark();
void DocumentBuilderInsertBreak();
void DocumentBuilderInsertElements();
void DocumentBuilderInsertHorizontalRule();
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
void ExtractTextOnly();
// TODO (std_string) : doesn't work due to dependebcy from Aspose.Barcode lib
//void GenerateACustomBarCodeImage.Run();
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
// TODO (std_string) : signatures isn't supported
//void SigningSignatureLine();
void UpdateContentControls();
// TODO(std_string) : absent documents
//void WorkingWithImportFormatOptions();
// TODO(std_string) : doesn't all work properly (in C# code)
void WorkingWithMarkdownFeatures();
void WorkingWithRevisions();
void WorkingWithSaveOptions();
void WriteAndFont();

// Fields
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
void InsertField();
void InsertFieldNone();
void InsertFormFields();
void InsertIncludeTextFieldWithoutDocumentBuilder();
void InsertMailMergeAddressBlockFieldUsingDOM();
void InsertMergeFieldUsingDOM();
void InsertNestedFields();
void InsertTOAFieldWithoutDocumentBuilder();
void RemoveField();
void RenameMergeFields();
void SpecifylocaleAtFieldlevel();
void UpdateDocFields();
void UseOfficeMathProperties();

// Images
void AddImageToEachPage();
void AddWatermark();
void CompressImages();
void ExtractImagesToFiles();
void InsertBarcodeImage();
void RemoveWatermark();

// Ranges
void RangesDeleteText();
void RangesGetText();

// Charts
void ChartNumberFormat();
void CreateChartUsingShape();
void CreateColumnChart();
void InsertAreaChart();
void InsertBubbleChart();
void InsertScatterChart();
void WorkingWithChartAxis();
void WorkWithChartDataLabel();
void WorkWithSingleChartDataPoint();
void WorkWithSingleChartSeries();

// Theme
void ManipulateThemeProperties();

// Node
void ExNode();

// Hyperlink
void ReplaceHyperlinks();

// Styles
void ChangeStyleOfTOCLevel();
void ChangeTOCTabStops();
void CopyStyles();
void ExtractContentBasedOnStyles();
void InsertStyleSeparator();

// Tables
void AddRemoveColumn();
void ApplyFormatting();
void ApplyStyle();
void AutoFitTableToContents();
void AutoFitTableToFixedColumnWidths();
void AutoFitTableToWindow();
// TODO(std_string) : isn't compiled due to using of DataTable
//void BuildTableFromDataTable();
void CloneTable();
void ExtractOrReplaceText();
void FindingIndex();
void InsertTableDirectly();
void InsertTableFromHtml();
void InsertTableUsingDocumentBuilder();
void JoiningAndSplittingTable();
void KeepTablesAndRowsBreaking();
void MergedCells();
void RepeatRowsOnSubsequentPages();
void SpecifyHeightAndWidth();
void TablePosition();

// Sections
void AddDeleteSection();
void AppendSectionContent();
void CloneSection();
void CopySection();
void DeleteHeaderFooterContent();
void DeleteSectionContent();
// TODO(std_string) : absent documents
//void ModifyPageSetupInAllSectionsOfDocument();
void SectionsAccessByIndex();

// StructuredDocumentTag 
void WorkingWithSDT();

// EndNote and Footnote
void WorkingWithFootnote();

// List
void WorkingWithList();

// Linked TextBoxes
void WorkingWithLinkedTextboxes();

// =====================================================
// Mail-Merge
// =====================================================
// TODO(std_string) : isn't compiled due to using of DataSet
//void ApplyCustomLogicToEmptyRegions();
void ExecuteArray();
// TODO(std_string) : isn't compiled due to using of OleDbXXX & DataTable
//void ExecuteWithRegionsDataTable();
void HandleMailMergeSwitches();
// TODO(std_string) : isn't compiled due to using of Linq2Xml
//void LINQtoXMLMailMerge();
// TODO(std_string) : isn't compiled due to using of DataTable
//void MailMergeAlternatingRows();
void MailMergeAndConditionalField();
void MailMergeCleanUp();
void MailMergeFormFields();
// TODO(std_string) : isn't compiled due to using of OleDbXXX & IDataReader
//void MailMergeImageFromBlob();
// TODO(std_string) : isn't compiled due to using of DataSet
//void MailMergeUsingMustacheSyntax.Run();
// TODO(std_string) : isn't compiled due to using of OleDbXXX & DataTable
//void MultipleDocsInMailMerge();
// TODO(std_string) : isn't compiled due to using of DataSet
//void NestedMailMerge();
void NestedMailMergeCustom();
// TODO(std_string) : OleDbXXX & DataTable
//void ProduceMultipleDocuments();
// TODO(std_string) : isn't compiled due to using of DataSet
//void RemoveEmptyRegions();
// TODO(std_string) : isn't compiled due to using of DataSet
//void RemoveRowsFromTable();
void SimpleMailMerge();
// TODO(std_string) : isn't compiled due to using of DataSet
//void XMLMailMerge();

// =====================================================
// Rendering and Printing
// =====================================================
void DocumentLayoutHelper();
void EmbeddedFontsInPDF();
void EmbeddingWindowsStandardFonts();
void EnumerateLayoutElements();
void HyphenateWordsOfLanguages();
// TODO (std_string) : saving into images don't work properly
//void ImageColorFilters();
void LoadHyphenationDictionaryForLanguage();
// TODO (std_string) : doesn't work due to using of GUI
//void PrintProgressDialog();
// TODO (std_string) : doesn't work due to using of GUI
//void Print_CachePrinterSettings();
// TODO (std_string) : doesn't work due to using of OleControl
//void ReadActiveXControlProperties();
void ReceiveNotificationsOfFont();
void RenderShape();
void ResourceSteamFontSource();
// TODO (std_string) : saving into images don't work properly
//void SaveAsMultipageTiff();
void SetFontSettings();
void SetFontsFoldersMultipleFolders();
void SetFontsFoldersSystemAndCustomFolder();
void SetHorizontalAndVerticalImageResolution();
void SetTrueTypeFontsFolder();
void SpecifyDefaultFontWhenRendering();
// TODO (std_string) : absent documents
//void WorkingWithFontResolution();
void WorkingWithFontSources();
void WorkingWithPdfSaveOptions();