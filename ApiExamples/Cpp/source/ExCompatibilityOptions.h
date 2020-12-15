#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Settings/CompatibilityOptions.h>
#include <Aspose.Words.Cpp/Model/Settings/MsWordVersion.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExCompatibilityOptions : public ApiExampleBase
{

public:
    void Tables()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2002);

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
        doc->Save(ArtifactsDir + u"CompatibilityOptions.Tables.docx");
    }

    void Breaks()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);

        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ApplyBreakingRules());
        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotUseEastAsianBreakRules());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ShowBreaksInFrames());
        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_SplitPgBreakAndParaMark());
        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UseAltKinsokuLineBreakRules());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UseWord97LineBreakRules());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc->Save(ArtifactsDir + u"CompatibilityOptions.Breaks.docx");
    }

    void Spacing()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);

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
        doc->Save(ArtifactsDir + u"CompatibilityOptions.Spacing.docx");
    }

    void WordPerfect()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);

        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_SuppressTopSpacingWP());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_TruncateFontHeightsLikeWP6());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_WPJustification());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_WPSpaceWidth());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_WrapTrailSpaces());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc->Save(ArtifactsDir + u"CompatibilityOptions.WordPerfect.docx");
    }

    void Alignment()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);

        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_CachedColBalance());
        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotVertAlignInTxbx());
        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_DoNotWrapTextWithPunct());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_NoTabHangInd());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc->Save(ArtifactsDir + u"CompatibilityOptions.Alignment.docx");
    }

    void Legacy()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);

        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_FootnoteLayoutLikeWW8());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_LineWrapLikeWord6());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_MWSmallCaps());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_ShapeLayoutLikeWW8());
        ASPOSE_ASSERT_EQ(false, compatibilityOptions->get_UICompat97To2003());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc->Save(ArtifactsDir + u"CompatibilityOptions.Legacy.docx");
    }

    void List()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);

        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UnderlineTabInNumList());
        ASPOSE_ASSERT_EQ(true, compatibilityOptions->get_UseNormalStyleForList());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc->Save(ArtifactsDir + u"CompatibilityOptions.List.docx");
    }

    void Misc()
    {
        auto doc = MakeObject<Document>();

        SharedPtr<CompatibilityOptions> compatibilityOptions = doc->get_CompatibilityOptions();
        compatibilityOptions->OptimizeFor(MsWordVersion::Word2000);

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
        doc->Save(ArtifactsDir + u"CompatibilityOptions.Misc.docx");
    }
};

} // namespace ApiExamples
