#include "ExUtilityClasses.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
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
    pageSetup->set_TopMargin(ConvertUtil::InchToPoint(1.0));
    pageSetup->set_BottomMargin(ConvertUtil::InchToPoint(2.0));
    pageSetup->set_LeftMargin(ConvertUtil::InchToPoint(2.5));
    pageSetup->set_RightMargin(ConvertUtil::InchToPoint(1.5));
    
    // An inch is 72 points.
    ASPOSE_ASSERT_EQ(ConvertUtil::InchToPoint(1), 72.0);
    ASPOSE_ASSERT_EQ(ConvertUtil::PointToInch(72), 1.0);
    
    // Add content to demonstrate the new margins.
    builder->Writeln(System::String::Format(u"This Text is {0} points/{1} inches from the left, ", pageSetup->get_LeftMargin(), ConvertUtil::PointToInch(pageSetup->get_LeftMargin())) + System::String::Format(u"{0} points/{1} inches from the right, ", pageSetup->get_RightMargin(), ConvertUtil::PointToInch(pageSetup->get_RightMargin())) + System::String::Format(u"{0} points/{1} inches from the top, ", pageSetup->get_TopMargin(), ConvertUtil::PointToInch(pageSetup->get_TopMargin())) + System::String::Format(u"and {0} points/{1} inches from the bottom of the page.", pageSetup->get_BottomMargin(), ConvertUtil::PointToInch(pageSetup->get_BottomMargin())));
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndInches.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"UtilityClasses.PointsAndInches.docx"));
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(pageSetup->get_TopMargin(), 72.0, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToInch(pageSetup->get_TopMargin()), 1.0, 0.01);
    ASSERT_NEAR(pageSetup->get_BottomMargin(), 144.0, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToInch(pageSetup->get_BottomMargin()), 2.0, 0.01);
    ASSERT_NEAR(pageSetup->get_LeftMargin(), 180.0, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToInch(pageSetup->get_LeftMargin()), 2.5, 0.01);
    ASSERT_NEAR(pageSetup->get_RightMargin(), 108.0, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToInch(pageSetup->get_RightMargin()), 1.5, 0.01);
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
    pageSetup->set_TopMargin(ConvertUtil::MillimeterToPoint(30));
    pageSetup->set_BottomMargin(ConvertUtil::MillimeterToPoint(50));
    pageSetup->set_LeftMargin(ConvertUtil::MillimeterToPoint(80));
    pageSetup->set_RightMargin(ConvertUtil::MillimeterToPoint(40));
    
    // A centimeter is approximately 28.3 points.
    ASSERT_NEAR(ConvertUtil::MillimeterToPoint(10), 28.34, 0.01);
    
    // Add content to demonstrate the new margins.
    builder->Writeln(System::String::Format(u"This Text is {0} points from the left, ", pageSetup->get_LeftMargin()) + System::String::Format(u"{0} points from the right, ", pageSetup->get_RightMargin()) + System::String::Format(u"{0} points from the top, ", pageSetup->get_TopMargin()) + System::String::Format(u"and {0} points from the bottom of the page.", pageSetup->get_BottomMargin()));
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndMillimeters.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"UtilityClasses.PointsAndMillimeters.docx"));
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(pageSetup->get_TopMargin(), 85.05, 0.01);
    ASSERT_NEAR(pageSetup->get_BottomMargin(), 141.75, 0.01);
    ASSERT_NEAR(pageSetup->get_LeftMargin(), 226.75, 0.01);
    ASSERT_NEAR(pageSetup->get_RightMargin(), 113.4, 0.01);
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
    pageSetup->set_TopMargin(ConvertUtil::PixelToPoint(static_cast<double>(100)));
    pageSetup->set_BottomMargin(ConvertUtil::PixelToPoint(static_cast<double>(200)));
    pageSetup->set_LeftMargin(ConvertUtil::PixelToPoint(static_cast<double>(225)));
    pageSetup->set_RightMargin(ConvertUtil::PixelToPoint(static_cast<double>(125)));
    
    // A pixel is 0.75 points.
    ASPOSE_ASSERT_EQ(ConvertUtil::PixelToPoint(static_cast<double>(1)), 0.75);
    ASPOSE_ASSERT_EQ(ConvertUtil::PointToPixel(0.75), 1.0);
    
    // The default DPI value used is 96.
    ASPOSE_ASSERT_EQ(ConvertUtil::PixelToPoint(static_cast<double>(1), static_cast<double>(96)), 0.75);
    
    // Add content to demonstrate the new margins.
    builder->Writeln(System::String::Format(u"This Text is {0} points/{1} pixels from the left, ", pageSetup->get_LeftMargin(), ConvertUtil::PointToPixel(pageSetup->get_LeftMargin())) + System::String::Format(u"{0} points/{1} pixels from the right, ", pageSetup->get_RightMargin(), ConvertUtil::PointToPixel(pageSetup->get_RightMargin())) + System::String::Format(u"{0} points/{1} pixels from the top, ", pageSetup->get_TopMargin(), ConvertUtil::PointToPixel(pageSetup->get_TopMargin())) + System::String::Format(u"and {0} points/{1} pixels from the bottom of the page.", pageSetup->get_BottomMargin(), ConvertUtil::PointToPixel(pageSetup->get_BottomMargin())));
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixels.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixels.docx"));
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(pageSetup->get_TopMargin(), 75.0, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(pageSetup->get_TopMargin()), 100.0, 0.01);
    ASSERT_NEAR(pageSetup->get_BottomMargin(), 150.0, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(pageSetup->get_BottomMargin()), 200.0, 0.01);
    ASSERT_NEAR(pageSetup->get_LeftMargin(), 168.75, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(pageSetup->get_LeftMargin()), 225.0, 0.01);
    ASSERT_NEAR(pageSetup->get_RightMargin(), 93.75, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(pageSetup->get_RightMargin()), 125.0, 0.01);
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
    pageSetup->set_TopMargin(ConvertUtil::PixelToPoint(static_cast<double>(100), myDpi));
    
    ASSERT_NEAR(pageSetup->get_TopMargin(), 37.5, 0.01);
    
    // At the default DPI of 96, a pixel is 0.75 points.
    ASPOSE_ASSERT_EQ(ConvertUtil::PixelToPoint(static_cast<double>(1)), 0.75);
    
    builder->Writeln(System::String::Format(u"This Text is {0} points/{1} ", pageSetup->get_TopMargin(), ConvertUtil::PointToPixel(pageSetup->get_TopMargin(), myDpi)) + System::String::Format(u"pixels (at a DPI of {0}) from the top of the page.", myDpi));
    
    // Set a new DPI and adjust the top margin value accordingly.
    const double newDpi = 300;
    pageSetup->set_TopMargin(ConvertUtil::PixelToNewDpi(pageSetup->get_TopMargin(), myDpi, newDpi));
    ASSERT_NEAR(pageSetup->get_TopMargin(), 59.0, 0.01);
    
    builder->Writeln(System::String::Format(u"At a DPI of {0}, the text is now {1} points/{2} ", newDpi, pageSetup->get_TopMargin(), ConvertUtil::PointToPixel(pageSetup->get_TopMargin(), myDpi)) + u"pixels from the top of the page.");
    
    doc->Save(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixelsDpi.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"UtilityClasses.PointsAndPixelsDpi.docx"));
    pageSetup = doc->get_FirstSection()->get_PageSetup();
    
    ASSERT_NEAR(pageSetup->get_TopMargin(), 59.0, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(pageSetup->get_TopMargin()), 78.66, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(pageSetup->get_TopMargin(), myDpi), 157.33, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(static_cast<double>(100)), 133.33, 0.01);
    ASSERT_NEAR(ConvertUtil::PointToPixel(static_cast<double>(100), myDpi), 266.66, 0.01);
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
