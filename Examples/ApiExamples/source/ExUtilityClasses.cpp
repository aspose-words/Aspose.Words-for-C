// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExUtilityClasses.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/ConvertUtil.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2171840234u, ::Aspose::Words::ApiExamples::ExUtilityClasses, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExUtilityClasses : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExUtilityClasses> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExUtilityClasses>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExUtilityClasses> ExUtilityClasses::s_instance;

} // namespace gtest_test

void ExUtilityClasses::PointsAndInches()
{
    //ExStart
    //ExFor:ConvertUtil
    //ExFor:ConvertUtil.PointToInch
    //ExFor:ConvertUtil.InchToPoint
    //ExSummary:Shows how to specify page properties in inches.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A section's "Page Setup" defines the size of the page margins in points.
    // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
    // such as inches when defining boundaries.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = builder->get_PageSetup();
    pageSetup->set_TopMargin(Aspose::Words::ConvertUtil::InchToPoint(1.0));
    pageSetup->set_BottomMargin(Aspose::Words::ConvertUtil::InchToPoint(2.0));
    pageSetup->set_LeftMargin(Aspose::Words::ConvertUtil::InchToPoint(2.5));
    pageSetup->set_RightMargin(Aspose::Words::ConvertUtil::InchToPoint(1.5));
    
    // An inch is 72 points.
    ASPOSE_ASSERT_EQ(72.0, Aspose::Words::ConvertUtil::InchToPoint(1));
    ASPOSE_ASSERT_EQ(1.0, Aspose::Words::ConvertUtil::PointToInch(72));
    
    // Add content to demonstrate the new margins.
    builder->Writeln(System::String::Format(u"This Text is {0} points/{1} inches from the left, ", pageSetup->get_LeftMargin(), Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_LeftMargin())) + System::String::Format(u"{0} points/{1} inches from the right, ", pageSetup->get_RightMargin(), Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_RightMargin())) + System::String::Format(u"{0} points/{1} inches from the top, ", pageSetup->get_TopMargin(), Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_TopMargin())) + System::String::Format(u"and {0} points/{1} inches from the bottom of the page.", pageSetup->get_BottomMargin(), Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_BottomMargin())));
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndInches.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"UtilityClasses.PointsAndInches.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(72.0, pageSetup->get_TopMargin(), 0.01);
    ASSERT_NEAR(1.0, Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_TopMargin()), 0.01);
    ASSERT_NEAR(144.0, pageSetup->get_BottomMargin(), 0.01);
    ASSERT_NEAR(2.0, Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_BottomMargin()), 0.01);
    ASSERT_NEAR(180.0, pageSetup->get_LeftMargin(), 0.01);
    ASSERT_NEAR(2.5, Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_LeftMargin()), 0.01);
    ASSERT_NEAR(108.0, pageSetup->get_RightMargin(), 0.01);
    ASSERT_NEAR(1.5, Aspose::Words::ConvertUtil::PointToInch(pageSetup->get_RightMargin()), 0.01);
}

namespace gtest_test
{

TEST_F(ExUtilityClasses, PointsAndInches)
{
    s_instance->PointsAndInches();
}

} // namespace gtest_test

void ExUtilityClasses::PointsAndMillimeters()
{
    //ExStart
    //ExFor:ConvertUtil.MillimeterToPoint
    //ExSummary:Shows how to specify page properties in millimeters.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A section's "Page Setup" defines the size of the page margins in points.
    // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
    // such as millimeters when defining boundaries.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = builder->get_PageSetup();
    pageSetup->set_TopMargin(Aspose::Words::ConvertUtil::MillimeterToPoint(30));
    pageSetup->set_BottomMargin(Aspose::Words::ConvertUtil::MillimeterToPoint(50));
    pageSetup->set_LeftMargin(Aspose::Words::ConvertUtil::MillimeterToPoint(80));
    pageSetup->set_RightMargin(Aspose::Words::ConvertUtil::MillimeterToPoint(40));
    
    // A centimeter is approximately 28.3 points.
    ASSERT_NEAR(28.34, Aspose::Words::ConvertUtil::MillimeterToPoint(10), 0.01);
    
    // Add content to demonstrate the new margins.
    builder->Writeln(System::String::Format(u"This Text is {0} points from the left, ", pageSetup->get_LeftMargin()) + System::String::Format(u"{0} points from the right, ", pageSetup->get_RightMargin()) + System::String::Format(u"{0} points from the top, ", pageSetup->get_TopMargin()) + System::String::Format(u"and {0} points from the bottom of the page.", pageSetup->get_BottomMargin()));
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndMillimeters.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"UtilityClasses.PointsAndMillimeters.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(85.05, pageSetup->get_TopMargin(), 0.01);
    ASSERT_NEAR(141.75, pageSetup->get_BottomMargin(), 0.01);
    ASSERT_NEAR(226.75, pageSetup->get_LeftMargin(), 0.01);
    ASSERT_NEAR(113.4, pageSetup->get_RightMargin(), 0.01);
}

namespace gtest_test
{

TEST_F(ExUtilityClasses, PointsAndMillimeters)
{
    s_instance->PointsAndMillimeters();
}

} // namespace gtest_test

void ExUtilityClasses::PointsAndPixels()
{
    //ExStart
    //ExFor:ConvertUtil.PixelToPoint(double)
    //ExFor:ConvertUtil.PointToPixel(double)
    //ExSummary:Shows how to specify page properties in pixels.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // A section's "Page Setup" defines the size of the page margins in points.
    // We can also use the "ConvertUtil" class to use a different measurement unit,
    // such as pixels when defining boundaries.
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = builder->get_PageSetup();
    pageSetup->set_TopMargin(Aspose::Words::ConvertUtil::PixelToPoint(100));
    pageSetup->set_BottomMargin(Aspose::Words::ConvertUtil::PixelToPoint(200));
    pageSetup->set_LeftMargin(Aspose::Words::ConvertUtil::PixelToPoint(225));
    pageSetup->set_RightMargin(Aspose::Words::ConvertUtil::PixelToPoint(125));
    
    // A pixel is 0.75 points.
    ASPOSE_ASSERT_EQ(0.75, Aspose::Words::ConvertUtil::PixelToPoint(1));
    ASPOSE_ASSERT_EQ(1.0, Aspose::Words::ConvertUtil::PointToPixel(0.75));
    
    // The default DPI value used is 96.
    ASPOSE_ASSERT_EQ(0.75, Aspose::Words::ConvertUtil::PixelToPoint(1, 96));
    
    // Add content to demonstrate the new margins.
    builder->Writeln(System::String::Format(u"This Text is {0} points/{1} pixels from the left, ", pageSetup->get_LeftMargin(), Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_LeftMargin())) + System::String::Format(u"{0} points/{1} pixels from the right, ", pageSetup->get_RightMargin(), Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_RightMargin())) + System::String::Format(u"{0} points/{1} pixels from the top, ", pageSetup->get_TopMargin(), Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_TopMargin())) + System::String::Format(u"and {0} points/{1} pixels from the bottom of the page.", pageSetup->get_BottomMargin(), Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_BottomMargin())));
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixels.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixels.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(75.0, pageSetup->get_TopMargin(), 0.01);
    ASSERT_NEAR(100.0, Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_TopMargin()), 0.01);
    ASSERT_NEAR(150.0, pageSetup->get_BottomMargin(), 0.01);
    ASSERT_NEAR(200.0, Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_BottomMargin()), 0.01);
    ASSERT_NEAR(168.75, pageSetup->get_LeftMargin(), 0.01);
    ASSERT_NEAR(225.0, Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_LeftMargin()), 0.01);
    ASSERT_NEAR(93.75, pageSetup->get_RightMargin(), 0.01);
    ASSERT_NEAR(125.0, Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_RightMargin()), 0.01);
}

namespace gtest_test
{

TEST_F(ExUtilityClasses, PointsAndPixels)
{
    s_instance->PointsAndPixels();
}

} // namespace gtest_test

void ExUtilityClasses::PointsAndPixelsDpi()
{
    //ExStart
    //ExFor:ConvertUtil.PixelToNewDpi
    //ExFor:ConvertUtil.PixelToPoint(double, double)
    //ExFor:ConvertUtil.PointToPixel(double, double)
    //ExSummary:Shows how to use convert points to pixels with default and custom resolution.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Define the size of the top margin of this section in pixels, according to a custom DPI.
    const double myDpi = 192;
    
    System::SharedPtr<Aspose::Words::PageSetup> pageSetup = builder->get_PageSetup();
    pageSetup->set_TopMargin(Aspose::Words::ConvertUtil::PixelToPoint(100, myDpi));
    
    ASSERT_NEAR(37.5, pageSetup->get_TopMargin(), 0.01);
    
    // At the default DPI of 96, a pixel is 0.75 points.
    ASPOSE_ASSERT_EQ(0.75, Aspose::Words::ConvertUtil::PixelToPoint(1));
    
    builder->Writeln(System::String::Format(u"This Text is {0} points/{1} ", pageSetup->get_TopMargin(), Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_TopMargin(), myDpi)) + System::String::Format(u"pixels (at a DPI of {0}) from the top of the page.", myDpi));
    
    // Set a new DPI and adjust the top margin value accordingly.
    const double newDpi = 300;
    pageSetup->set_TopMargin(Aspose::Words::ConvertUtil::PixelToNewDpi(pageSetup->get_TopMargin(), myDpi, newDpi));
    ASSERT_NEAR(59.0, pageSetup->get_TopMargin(), 0.01);
    
    builder->Writeln(System::String::Format(u"At a DPI of {0}, the text is now {1} points/{2} ", newDpi, pageSetup->get_TopMargin(), Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_TopMargin(), myDpi)) + u"pixels from the top of the page.");
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixelsDpi.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixelsDpi.docx");
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(59.0, pageSetup->get_TopMargin(), 0.01);
    ASSERT_NEAR(78.66, Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_TopMargin()), 0.01);
    ASSERT_NEAR(157.33, Aspose::Words::ConvertUtil::PointToPixel(pageSetup->get_TopMargin(), myDpi), 0.01);
    ASSERT_NEAR(133.33, Aspose::Words::ConvertUtil::PointToPixel(100), 0.01);
    ASSERT_NEAR(266.66, Aspose::Words::ConvertUtil::PointToPixel(100, myDpi), 0.01);
}

namespace gtest_test
{

TEST_F(ExUtilityClasses, PointsAndPixelsDpi)
{
    s_instance->PointsAndPixelsDpi();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
