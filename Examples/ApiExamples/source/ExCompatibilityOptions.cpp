#include "ExCompatibilityOptions.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/enumerator_adapter.h>
#include <system/console.h>
#include <system/collections/list.h>
#include <iostream>
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
    doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2010);
    std::cout << "\nOptimized for Word 2010:" << std::endl;
    PrintCompatibilityOptions(options);
    
    doc->get_CompatibilityOptions()->OptimizeFor(MsWordVersion::Word2000);
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2002);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_AdjustLineHeightInTable(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_AlignTablesRowByRow(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_AllowSpaceOfSameStyleInTable(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotAutofitConstrainedTables(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotBreakConstrainedForcedTable(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotBreakWrappedTables(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotSnapToGridInCell(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotUseHTMLParagraphAutoSpacing(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotVertAlignCellWithSp(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_ForgetLastTabAlignment(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_GrowAutofit(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_LayoutRawTableWidth(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_LayoutTableRowsApart(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_NoColumnBalance(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_OverrideTableStyleFontSizeAndJustification(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseSingleBorderforContiguousCells(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseWord2002TableStyleRules(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseWord2010TableStyleRules(), false);
    
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_ApplyBreakingRules(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotUseEastAsianBreakRules(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_ShowBreaksInFrames(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SplitPgBreakAndParaMark(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseAltKinsokuLineBreakRules(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseWord97LineBreakRules(), false);
    
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_AutoSpaceLikeWord95(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DisplayHangulFixedWidth(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_NoExtraLineSpacing(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_NoLeading(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_NoSpaceRaiseLower(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SpaceForUL(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SpacingInWholePoints(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SuppressBottomSpacing(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SuppressSpBfAfterPgBrk(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SuppressSpacingAtTopOfPage(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SuppressTopSpacing(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UlTrailSpace(), false);
    
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SuppressTopSpacingWP(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_TruncateFontHeightsLikeWP6(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_WPJustification(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_WPSpaceWidth(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_WrapTrailSpaces(), false);
    
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_CachedColBalance(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotVertAlignInTxbx(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotWrapTextWithPunct(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_NoTabHangInd(), false);
    
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_FootnoteLayoutLikeWW8(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_LineWrapLikeWord6(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_MWSmallCaps(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_ShapeLayoutLikeWW8(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UICompat97To2003(), false);
    
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UnderlineTabInNumList(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseNormalStyleForList(), true);
    
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
    compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);
    
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_BalanceSingleByteDoubleByteWidth(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_ConvMailMergeEsc(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotExpandShiftReturn(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotLeaveBackslashAlone(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotSuppressParagraphBorders(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_DoNotUseIndentAsNumberingTabStop(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_PrintBodyTextBeforeHeader(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_PrintColBlack(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SelectFldWithFirstOrLastChar(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SubFontBySize(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_SwapBordersFacingPgs(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_TransparentMetafiles(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseAnsiKerningPairs(), true);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UseFELayout(), false);
    ASPOSE_ASSERT_EQ(compatibilityOptions->get_UsePrinterMetrics(), false);
    
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
