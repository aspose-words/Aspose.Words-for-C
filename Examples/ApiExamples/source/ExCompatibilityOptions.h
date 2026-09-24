#pragma once

#include <system/string.h>
#include <system/collections/ilist.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Settings;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExCompatibilityOptions : public ApiExampleBase
{
    typedef ExCompatibilityOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    //ExStart
    //ExFor:Compatibility
    //ExFor:CompatibilityOptions
    //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
    //ExFor:Document.CompatibilityOptions
    //ExFor:MsWordVersion
    //ExFor:CompatibilityOptions.AdjustLineHeightInTable
    //ExFor:CompatibilityOptions.AlignTablesRowByRow
    //ExFor:CompatibilityOptions.AllowSpaceOfSameStyleInTable
    //ExFor:CompatibilityOptions.ApplyBreakingRules
    //ExFor:CompatibilityOptions.AutofitToFirstFixedWidthCell
    //ExFor:CompatibilityOptions.AutoSpaceLikeWord95
    //ExFor:CompatibilityOptions.BalanceSingleByteDoubleByteWidth
    //ExFor:CompatibilityOptions.CachedColBalance
    //ExFor:CompatibilityOptions.ConvMailMergeEsc
    //ExFor:CompatibilityOptions.DisableOpenTypeFontFormattingFeatures
    //ExFor:CompatibilityOptions.DisplayHangulFixedWidth
    //ExFor:CompatibilityOptions.DoNotAutofitConstrainedTables
    //ExFor:CompatibilityOptions.DoNotBreakConstrainedForcedTable
    //ExFor:CompatibilityOptions.DoNotBreakWrappedTables
    //ExFor:CompatibilityOptions.DoNotExpandShiftReturn
    //ExFor:CompatibilityOptions.DoNotLeaveBackslashAlone
    //ExFor:CompatibilityOptions.DoNotSnapToGridInCell
    //ExFor:CompatibilityOptions.DoNotSuppressIndentation
    //ExFor:CompatibilityOptions.DoNotSuppressParagraphBorders
    //ExFor:CompatibilityOptions.DoNotUseEastAsianBreakRules
    //ExFor:CompatibilityOptions.DoNotUseHTMLParagraphAutoSpacing
    //ExFor:CompatibilityOptions.DoNotUseIndentAsNumberingTabStop
    //ExFor:CompatibilityOptions.DoNotVertAlignCellWithSp
    //ExFor:CompatibilityOptions.DoNotVertAlignInTxbx
    //ExFor:CompatibilityOptions.DoNotWrapTextWithPunct
    //ExFor:CompatibilityOptions.FootnoteLayoutLikeWW8
    //ExFor:CompatibilityOptions.ForgetLastTabAlignment
    //ExFor:CompatibilityOptions.GrowAutofit
    //ExFor:CompatibilityOptions.LayoutRawTableWidth
    //ExFor:CompatibilityOptions.LayoutTableRowsApart
    //ExFor:CompatibilityOptions.LineWrapLikeWord6
    //ExFor:CompatibilityOptions.MWSmallCaps
    //ExFor:CompatibilityOptions.NoColumnBalance
    //ExFor:CompatibilityOptions.NoExtraLineSpacing
    //ExFor:CompatibilityOptions.NoLeading
    //ExFor:CompatibilityOptions.NoSpaceRaiseLower
    //ExFor:CompatibilityOptions.NoTabHangInd
    //ExFor:CompatibilityOptions.OverrideTableStyleFontSizeAndJustification
    //ExFor:CompatibilityOptions.PrintBodyTextBeforeHeader
    //ExFor:CompatibilityOptions.PrintColBlack
    //ExFor:CompatibilityOptions.SelectFldWithFirstOrLastChar
    //ExFor:CompatibilityOptions.ShapeLayoutLikeWW8
    //ExFor:CompatibilityOptions.ShowBreaksInFrames
    //ExFor:CompatibilityOptions.SpaceForUL
    //ExFor:CompatibilityOptions.SpacingInWholePoints
    //ExFor:CompatibilityOptions.SplitPgBreakAndParaMark
    //ExFor:CompatibilityOptions.SubFontBySize
    //ExFor:CompatibilityOptions.SuppressBottomSpacing
    //ExFor:CompatibilityOptions.SuppressSpacingAtTopOfPage
    //ExFor:CompatibilityOptions.SuppressSpBfAfterPgBrk
    //ExFor:CompatibilityOptions.SuppressTopSpacing
    //ExFor:CompatibilityOptions.SuppressTopSpacingWP
    //ExFor:CompatibilityOptions.SwapBordersFacingPgs
    //ExFor:CompatibilityOptions.SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning
    //ExFor:CompatibilityOptions.TransparentMetafiles
    //ExFor:CompatibilityOptions.TruncateFontHeightsLikeWP6
    //ExFor:CompatibilityOptions.UICompat97To2003
    //ExFor:CompatibilityOptions.UlTrailSpace
    //ExFor:CompatibilityOptions.UnderlineTabInNumList
    //ExFor:CompatibilityOptions.UseAltKinsokuLineBreakRules
    //ExFor:CompatibilityOptions.UseAnsiKerningPairs
    //ExFor:CompatibilityOptions.UseFELayout
    //ExFor:CompatibilityOptions.UseNormalStyleForList
    //ExFor:CompatibilityOptions.UsePrinterMetrics
    //ExFor:CompatibilityOptions.UseSingleBorderforContiguousCells
    //ExFor:CompatibilityOptions.UseWord2002TableStyleRules
    //ExFor:CompatibilityOptions.UseWord2010TableStyleRules
    //ExFor:CompatibilityOptions.UseWord97LineBreakRules
    //ExFor:CompatibilityOptions.WPJustification
    //ExFor:CompatibilityOptions.WPSpaceWidth
    //ExFor:CompatibilityOptions.WrapTrailSpaces
    //ExSummary:Shows how to optimize the document for different versions of Microsoft Word.
    void OptimizeFor();
    //ExEnd
    void Tables();
    void Breaks();
    void Spacing();
    void WordPerfect();
    void Alignment();
    void Legacy();
    void List();
    void Misc();
    
protected:

    /// <summary>
    /// Groups all flags in a document's compatibility options object by state, then prints each group.
    /// </summary>
    static void PrintCompatibilityOptions(System::SharedPtr<Aspose::Words::Settings::CompatibilityOptions> options);
    static void AddOptionName(bool option, System::String optionName, System::SharedPtr<System::Collections::Generic::IList<System::String>> enabledOptions, System::SharedPtr<System::Collections::Generic::IList<System::String>> disabledOptions);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


