// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCompatibilityOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/enumerator_adapter.h>
#include <system/collections/list.h>
#include <iostream>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Settings;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3123663518u, ::Aspose::Words::ApiExamples::ExCompatibilityOptions, ThisTypeBaseTypesInfo);

void ExCompatibilityOptions::PrintCompatibilityOptions(System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> options)
{
    System::SharedPtr<System::Collections::Generic::IList<System::String>> enabledOptions = System::MakeObject<System::Collections::Generic::List<System::String>>();
    System::SharedPtr<System::Collections::Generic::IList<System::String>> disabledOptions = System::MakeObject<System::Collections::Generic::List<System::String>>();
    AddOptionName(options->get_AdjustLineHeightInTable(), u"AdjustLineHeightInTable", enabledOptions, disabledOptions);
    AddOptionName(options->get_AlignTablesRowByRow(), u"AlignTablesRowByRow", enabledOptions, disabledOptions);
    AddOptionName(options->get_AllowSpaceOfSameStyleInTable(), u"AllowSpaceOfSameStyleInTable", enabledOptions, disabledOptions);
    AddOptionName(options->get_ApplyBreakingRules(), u"ApplyBreakingRules", enabledOptions, disabledOptions);
    AddOptionName(options->get_AutoSpaceLikeWord95(), u"AutoSpaceLikeWord95", enabledOptions, disabledOptions);
    AddOptionName(options->get_AutofitToFirstFixedWidthCell(), u"AutofitToFirstFixedWidthCell", enabledOptions, disabledOptions);
    AddOptionName(options->get_BalanceSingleByteDoubleByteWidth(), u"BalanceSingleByteDoubleByteWidth", enabledOptions, disabledOptions);
    AddOptionName(options->get_CachedColBalance(), u"CachedColBalance", enabledOptions, disabledOptions);
    AddOptionName(options->get_ConvMailMergeEsc(), u"ConvMailMergeEsc", enabledOptions, disabledOptions);
    AddOptionName(options->get_DisableOpenTypeFontFormattingFeatures(), u"DisableOpenTypeFontFormattingFeatures", enabledOptions, disabledOptions);
    AddOptionName(options->get_DisplayHangulFixedWidth(), u"DisplayHangulFixedWidth", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotAutofitConstrainedTables(), u"DoNotAutofitConstrainedTables", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotBreakConstrainedForcedTable(), u"DoNotBreakConstrainedForcedTable", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotBreakWrappedTables(), u"DoNotBreakWrappedTables", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotExpandShiftReturn(), u"DoNotExpandShiftReturn", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotLeaveBackslashAlone(), u"DoNotLeaveBackslashAlone", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotSnapToGridInCell(), u"DoNotSnapToGridInCell", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotSuppressIndentation(), u"DoNotSnapToGridInCell", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotSuppressParagraphBorders(), u"DoNotSuppressParagraphBorders", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotUseEastAsianBreakRules(), u"DoNotUseEastAsianBreakRules", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotUseHTMLParagraphAutoSpacing(), u"DoNotUseHTMLParagraphAutoSpacing", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotUseIndentAsNumberingTabStop(), u"DoNotUseIndentAsNumberingTabStop", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotVertAlignCellWithSp(), u"DoNotVertAlignCellWithSp", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotVertAlignInTxbx(), u"DoNotVertAlignInTxbx", enabledOptions, disabledOptions);
    AddOptionName(options->get_DoNotWrapTextWithPunct(), u"DoNotWrapTextWithPunct", enabledOptions, disabledOptions);
    AddOptionName(options->get_FootnoteLayoutLikeWW8(), u"FootnoteLayoutLikeWW8", enabledOptions, disabledOptions);
    AddOptionName(options->get_ForgetLastTabAlignment(), u"ForgetLastTabAlignment", enabledOptions, disabledOptions);
    AddOptionName(options->get_GrowAutofit(), u"GrowAutofit", enabledOptions, disabledOptions);
    AddOptionName(options->get_LayoutRawTableWidth(), u"LayoutRawTableWidth", enabledOptions, disabledOptions);
    AddOptionName(options->get_LayoutTableRowsApart(), u"LayoutTableRowsApart", enabledOptions, disabledOptions);
    AddOptionName(options->get_LineWrapLikeWord6(), u"LineWrapLikeWord6", enabledOptions, disabledOptions);
    AddOptionName(options->get_MWSmallCaps(), u"MWSmallCaps", enabledOptions, disabledOptions);
    AddOptionName(options->get_NoColumnBalance(), u"NoColumnBalance", enabledOptions, disabledOptions);
    AddOptionName(options->get_NoExtraLineSpacing(), u"NoExtraLineSpacing", enabledOptions, disabledOptions);
    AddOptionName(options->get_NoLeading(), u"NoLeading", enabledOptions, disabledOptions);
    AddOptionName(options->get_NoSpaceRaiseLower(), u"NoSpaceRaiseLower", enabledOptions, disabledOptions);
    AddOptionName(options->get_NoTabHangInd(), u"NoTabHangInd", enabledOptions, disabledOptions);
    AddOptionName(options->get_OverrideTableStyleFontSizeAndJustification(), u"OverrideTableStyleFontSizeAndJustification", enabledOptions, disabledOptions);
    AddOptionName(options->get_PrintBodyTextBeforeHeader(), u"PrintBodyTextBeforeHeader", enabledOptions, disabledOptions);
    AddOptionName(options->get_PrintColBlack(), u"PrintColBlack", enabledOptions, disabledOptions);
    AddOptionName(options->get_SelectFldWithFirstOrLastChar(), u"SelectFldWithFirstOrLastChar", enabledOptions, disabledOptions);
    AddOptionName(options->get_ShapeLayoutLikeWW8(), u"ShapeLayoutLikeWW8", enabledOptions, disabledOptions);
    AddOptionName(options->get_ShowBreaksInFrames(), u"ShowBreaksInFrames", enabledOptions, disabledOptions);
    AddOptionName(options->get_SpaceForUL(), u"SpaceForUL", enabledOptions, disabledOptions);
    AddOptionName(options->get_SpacingInWholePoints(), u"SpacingInWholePoints", enabledOptions, disabledOptions);
    AddOptionName(options->get_SplitPgBreakAndParaMark(), u"SplitPgBreakAndParaMark", enabledOptions, disabledOptions);
    AddOptionName(options->get_SubFontBySize(), u"SubFontBySize", enabledOptions, disabledOptions);
    AddOptionName(options->get_SuppressBottomSpacing(), u"SuppressBottomSpacing", enabledOptions, disabledOptions);
    AddOptionName(options->get_SuppressSpBfAfterPgBrk(), u"SuppressSpBfAfterPgBrk", enabledOptions, disabledOptions);
    AddOptionName(options->get_SuppressSpacingAtTopOfPage(), u"SuppressSpacingAtTopOfPage", enabledOptions, disabledOptions);
    AddOptionName(options->get_SuppressTopSpacing(), u"SuppressTopSpacing", enabledOptions, disabledOptions);
    AddOptionName(options->get_SuppressTopSpacingWP(), u"SuppressTopSpacingWP", enabledOptions, disabledOptions);
    AddOptionName(options->get_SwapBordersFacingPgs(), u"SwapBordersFacingPgs", enabledOptions, disabledOptions);
    AddOptionName(options->get_SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning(), u"SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning", enabledOptions, disabledOptions);
    AddOptionName(options->get_TransparentMetafiles(), u"TransparentMetafiles", enabledOptions, disabledOptions);
    AddOptionName(options->get_TruncateFontHeightsLikeWP6(), u"TruncateFontHeightsLikeWP6", enabledOptions, disabledOptions);
    AddOptionName(options->get_UICompat97To2003(), u"UICompat97To2003", enabledOptions, disabledOptions);
    AddOptionName(options->get_UlTrailSpace(), u"UlTrailSpace", enabledOptions, disabledOptions);
    AddOptionName(options->get_UnderlineTabInNumList(), u"UnderlineTabInNumList", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseAltKinsokuLineBreakRules(), u"UseAltKinsokuLineBreakRules", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseAnsiKerningPairs(), u"UseAnsiKerningPairs", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseFELayout(), u"UseFELayout", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseNormalStyleForList(), u"UseNormalStyleForList", enabledOptions, disabledOptions);
    AddOptionName(options->get_UsePrinterMetrics(), u"UsePrinterMetrics", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseSingleBorderforContiguousCells(), u"UseSingleBorderforContiguousCells", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseWord2002TableStyleRules(), u"UseWord2002TableStyleRules", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseWord2010TableStyleRules(), u"UseWord2010TableStyleRules", enabledOptions, disabledOptions);
    AddOptionName(options->get_UseWord97LineBreakRules(), u"UseWord97LineBreakRules", enabledOptions, disabledOptions);
    AddOptionName(options->get_WPJustification(), u"WPJustification", enabledOptions, disabledOptions);
    AddOptionName(options->get_WPSpaceWidth(), u"WPSpaceWidth", enabledOptions, disabledOptions);
    AddOptionName(options->get_WrapTrailSpaces(), u"WrapTrailSpaces", enabledOptions, disabledOptions);
    std::cout << "\tEnabled options:" << std::endl;
    for (auto&& optionName : System::IterateOver(enabledOptions))
    {
        std::cout << System::String::Format(u"\t\t{0}", optionName) << std::endl;
    }
    std::cout << "\tDisabled options:" << std::endl;
    for (auto&& optionName : System::IterateOver(disabledOptions))
    {
        std::cout << System::String::Format(u"\t\t{0}", optionName) << std::endl;
    }
}

void ExCompatibilityOptions::AddOptionName(bool option, System::String optionName, System::SharedPtr<System::Collections::Generic::IList<System::String>> enabledOptions, System::SharedPtr<System::Collections::Generic::IList<System::String>> disabledOptions)
{
    if (option)
    {
        enabledOptions->Add(optionName);
    }
    else
    {
        disabledOptions->Add(optionName);
    }
}


namespace gtest_test
{

class ExCompatibilityOptions : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExCompatibilityOptions> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExCompatibilityOptions>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExCompatibilityOptions> ExCompatibilityOptions::s_instance;

} // namespace gtest_test

void ExCompatibilityOptions::OptimizeFor()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // This object contains an extensive list of flags unique to each document
    // that allow us to facilitate backward compatibility with older versions of Microsoft Word.
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> options = doc->get_CompatibilityOptions();
    
    // Print the default settings for a blank document.
    std::cout << "\nDefault optimization settings:" << std::endl;
    PrintCompatibilityOptions(options);
    
    // We can access these settings in Microsoft Word via "File" -> "Options" -> "Advanced" -> "Compatibility options for...".
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.OptimizeFor.DefaultSettings.docx");
    
    // We can use the OptimizeFor method to ensure optimal compatibility with a specific Microsoft Word version.
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2010);
    std::cout << "\nOptimized for Word 2010:" << std::endl;
    PrintCompatibilityOptions(options);
    
    doc->get_CompatibilityOptions()->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    std::cout << "\nOptimized for Word 2000:" << std::endl;
    PrintCompatibilityOptions(options);
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, OptimizeFor)
{
    s_instance->OptimizeFor();
}

} // namespace gtest_test

void ExCompatibilityOptions::Tables()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2002);
    
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_AdjustLineHeightInTable());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_AlignTablesRowByRow());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_AllowSpaceOfSameStyleInTable());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotAutofitConstrainedTables());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotBreakConstrainedForcedTable());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_DoNotBreakWrappedTables());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_DoNotSnapToGridInCell());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_DoNotUseHTMLParagraphAutoSpacing());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotVertAlignCellWithSp());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ForgetLastTabAlignment());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_GrowAutofit());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_LayoutRawTableWidth());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_LayoutTableRowsApart());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_NoColumnBalance());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_OverrideTableStyleFontSizeAndJustification());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UseSingleBorderforContiguousCells());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UseWord2002TableStyleRules());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UseWord2010TableStyleRules());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.Tables.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, Tables)
{
    s_instance->Tables();
}

} // namespace gtest_test

void ExCompatibilityOptions::Breaks()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ApplyBreakingRules());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotUseEastAsianBreakRules());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ShowBreaksInFrames());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_SplitPgBreakAndParaMark());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UseAltKinsokuLineBreakRules());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UseWord97LineBreakRules());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.Breaks.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, Breaks)
{
    s_instance->Breaks();
}

} // namespace gtest_test

void ExCompatibilityOptions::Spacing()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_AutoSpaceLikeWord95());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DisplayHangulFixedWidth());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_NoExtraLineSpacing());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_NoLeading());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_NoSpaceRaiseLower());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SpaceForUL());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SpacingInWholePoints());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SuppressBottomSpacing());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SuppressSpBfAfterPgBrk());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SuppressSpacingAtTopOfPage());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SuppressTopSpacing());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UlTrailSpace());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.Spacing.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, Spacing)
{
    s_instance->Spacing();
}

} // namespace gtest_test

void ExCompatibilityOptions::WordPerfect()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SuppressTopSpacingWP());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_TruncateFontHeightsLikeWP6());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_WPJustification());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_WPSpaceWidth());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_WrapTrailSpaces());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.WordPerfect.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, WordPerfect)
{
    s_instance->WordPerfect();
}

} // namespace gtest_test

void ExCompatibilityOptions::Alignment()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_CachedColBalance());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotVertAlignInTxbx());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotWrapTextWithPunct());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_NoTabHangInd());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.Alignment.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, Alignment)
{
    s_instance->Alignment();
}

} // namespace gtest_test

void ExCompatibilityOptions::Legacy()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_FootnoteLayoutLikeWW8());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_LineWrapLikeWord6());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_MWSmallCaps());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ShapeLayoutLikeWW8());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UICompat97To2003());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.Legacy.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, Legacy)
{
    s_instance->Legacy();
}

} // namespace gtest_test

void ExCompatibilityOptions::List()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UnderlineTabInNumList());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UseNormalStyleForList());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.List.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, List)
{
    s_instance->List();
}

} // namespace gtest_test

void ExCompatibilityOptions::Misc()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
    compatibilityOptions->OptimizeFor(Aspose::Words::Settings::MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_BalanceSingleByteDoubleByteWidth());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ConvMailMergeEsc());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_DoNotExpandShiftReturn());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_DoNotLeaveBackslashAlone());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_DoNotSuppressParagraphBorders());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotUseIndentAsNumberingTabStop());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_PrintBodyTextBeforeHeader());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_PrintColBlack());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_SelectFldWithFirstOrLastChar());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SubFontBySize());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SwapBordersFacingPgs());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_TransparentMetafiles());
    ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UseAnsiKerningPairs());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UseFELayout());
    ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UsePrinterMetrics());
    
    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc->Save(get_ArtifactsDir() + u"CompatibilityOptions.Misc.docx");
}

namespace gtest_test
{

TEST_F(ExCompatibilityOptions, Misc)
{
    s_instance->Misc();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
