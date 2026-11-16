#include "ExDocumentProperties.h"

#include <testing/test_predicates.h>
#include <system/timespan.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/setters_wrap.h>
#include <system/object_ext.h>
#include <system/io/file.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/convert.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <system/collections/icollection.h>
#include <system/array.h>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Settings/WriteProtection.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Properties/PropertyType.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentSecurity.h>
#include <Aspose.Words.Cpp/Model/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Model/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldDocProperty.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkStart.h>
#include <Aspose.Words.Cpp/Model/Bookmarks/BookmarkEnd.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>

#include "TestUtil.h"
#include "DocumentHelper.h"


using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Properties;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1709644714u, ::Aspose::Words::ApiExamples::ExDocumentProperties::LineCounter, ThisTypeBaseTypesInfo);

ExDocumentProperties::LineCounter::LineCounter(System::SharedPtr<Aspose::Words::Document> doc) : mLineCount(0)
    , mScanningLineForRealText(false)
{
    mLayoutEnumerator = System::MakeObject<Aspose::Words::Layout::LayoutEnumerator>(doc);
    
    CountLines();
}

int32_t ExDocumentProperties::LineCounter::GetLineCount()
{
    return mLineCount;
}

void ExDocumentProperties::LineCounter::CountLines()
{
    do
    {
        if (mLayoutEnumerator->get_Type() == LayoutEntityType::Line)
        {
            mScanningLineForRealText = true;
        }
        
        if (mLayoutEnumerator->MoveFirstChild())
        {
            if (mScanningLineForRealText && mLayoutEnumerator->get_Kind().StartsWith(u"TEXT"))
            {
                mLineCount++;
                mScanningLineForRealText = false;
            }
            CountLines();
            mLayoutEnumerator->MoveParent();
        }
    } while (mLayoutEnumerator->MoveNext());
}


RTTI_INFO_IMPL_HASH(1422654878u, ::Aspose::Words::ApiExamples::ExDocumentProperties, ThisTypeBaseTypesInfo);

void ExDocumentProperties::TestContent(System::SharedPtr<Aspose::Words::Document> doc)
{
    System::SharedPtr<Aspose::Words::Properties::BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
    
    ASSERT_EQ(properties->get_Pages(), 6);
    
    ASSERT_EQ(properties->get_Words(), 1035);
    ASSERT_EQ(properties->get_Characters(), 6026);
    ASSERT_EQ(properties->get_CharactersWithSpaces(), 7041);
    ASSERT_EQ(properties->get_Lines(), 142);
    ASSERT_EQ(properties->get_Paragraphs(), 29);
    ASSERT_NEAR(properties->get_Bytes(), 15500, 200);
    ASSERT_EQ(properties->get_Template(), get_MyDir().Replace(u"\\\\", u"\\") + u"Business brochure.dotx");
    ASSERT_EQ(properties->get_ContentStatus(), u"Draft");
    ASSERT_EQ(properties->get_ContentType(), System::String::Empty);
    ASSERT_FALSE(properties->get_LinksUpToDate());
}

namespace gtest_test
{

class ExDocumentProperties : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentProperties> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDocumentProperties>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentProperties> ExDocumentProperties::s_instance;

} // namespace gtest_test

void ExDocumentProperties::BuiltIn()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties
    //ExFor:Document.BuiltInDocumentProperties
    //ExFor:Document.CustomDocumentProperties
    //ExFor:DocumentProperty
    //ExFor:DocumentProperty.Name
    //ExFor:DocumentProperty.Value
    //ExFor:DocumentProperty.Type
    //ExSummary:Shows how to work with built-in document properties.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Properties.docx"));
    
    // The "Document" object contains some of its metadata in its members.
    std::cout << System::String::Format(u"Document filename:\n\t \"{0}\"", doc->get_OriginalFileName()) << std::endl;
    
    // The document also stores metadata in its built-in properties.
    // Each built-in property is a member of the document's "BuiltInDocumentProperties" object.
    std::cout << "Built-in Properties:" << std::endl;
    for (auto&& docProperty : System::IterateOver(doc->get_BuiltInDocumentProperties()))
    {
        std::cout << docProperty->get_Name() << std::endl;
        std::cout << System::String::Format(u"\tType:\t{0}", docProperty->get_Type()) << std::endl;
        
        // Some properties may store multiple values.
        if (System::ObjectExt::Is<System::Collections::Generic::ICollection<System::SharedPtr<System::Object>>>(docProperty->get_Value()))
        {
            for (auto&& value : System::IterateOver(System::AsCast<System::Collections::Generic::ICollection<System::SharedPtr<System::Object>>>(docProperty->get_Value())))
            {
                std::cout << System::String::Format(u"\tValue:\t\"{0}\"", value) << std::endl;
            }
        }
        else
        {
            std::cout << System::String::Format(u"\tValue:\t\"{0}\"", docProperty->get_Value()) << std::endl;
        }
    }
    //ExEnd
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Count(), 31);
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, BuiltIn)
{
    s_instance->BuiltIn();
}

} // namespace gtest_test

void ExDocumentProperties::Custom()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.Item(String)
    //ExFor:CustomDocumentProperties
    //ExFor:DocumentProperty.ToString
    //ExFor:DocumentPropertyCollection.Count
    //ExFor:DocumentPropertyCollection.Item(int)
    //ExSummary:Shows how to work with custom document properties.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Properties.docx"));
    
    // Every document contains a collection of custom properties, which, like the built-in properties, are key-value pairs.
    // The document has a fixed list of built-in properties. The user creates all of the custom properties.
    ASSERT_EQ(System::ObjectExt::ToString(doc->get_CustomDocumentProperties()->idx_get(u"CustomProperty")), u"Value of custom document property");
    
    doc->get_CustomDocumentProperties()->Add(u"CustomProperty2", System::String(u"Value of custom document property #2"));
    
    std::cout << "Custom Properties:" << std::endl;
    for (auto&& customDocumentProperty : System::IterateOver(doc->get_CustomDocumentProperties()))
    {
        std::cout << customDocumentProperty->get_Name() << std::endl;
        std::cout << System::String::Format(u"\tType:\t{0}", customDocumentProperty->get_Type()) << std::endl;
        std::cout << System::String::Format(u"\tValue:\t\"{0}\"", customDocumentProperty->get_Value()) << std::endl;
    }
    //ExEnd
    
    ASSERT_EQ(doc->get_CustomDocumentProperties()->get_Count(), 2);
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, Custom)
{
    s_instance->Custom();
}

} // namespace gtest_test

void ExDocumentProperties::Description()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.Author
    //ExFor:BuiltInDocumentProperties.Category
    //ExFor:BuiltInDocumentProperties.Comments
    //ExFor:BuiltInDocumentProperties.Keywords
    //ExFor:BuiltInDocumentProperties.Subject
    //ExFor:BuiltInDocumentProperties.Title
    //ExSummary:Shows how to work with built-in document properties in the "Description" category.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    System::SharedPtr<Aspose::Words::Properties::BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
    
    // Below are four built-in document properties that have fields that can display their values in the document body.
    // 1 -  "Author" property, which we can display using an AUTHOR field:
    properties->set_Author(u"John Doe");
    builder->Write(u"Author:\t");
    builder->InsertField(FieldType::FieldAuthor, true);
    
    // 2 -  "Title" property, which we can display using a TITLE field:
    properties->set_Title(u"John's Document");
    builder->Write(u"\nDoc title:\t");
    builder->InsertField(FieldType::FieldTitle, true);
    
    // 3 -  "Subject" property, which we can display using a SUBJECT field:
    properties->set_Subject(u"My subject");
    builder->Write(u"\nSubject:\t");
    builder->InsertField(FieldType::FieldSubject, true);
    
    // 4 -  "Comments" property, which we can display using a COMMENTS field:
    properties->set_Comments(System::String::Format(u"This is {0}'s document about {1}", properties->get_Author(), properties->get_Subject()));
    builder->Write(u"\nComments:\t\"");
    builder->InsertField(FieldType::FieldComments, true);
    builder->Write(u"\"");
    
    // The "Category" built-in property does not have a field that can display its value.
    properties->set_Category(u"My category");
    
    // We can set multiple keywords for a document by separating the string value of the "Keywords" property with semicolons.
    properties->set_Keywords(u"Tag 1; Tag 2; Tag 3");
    
    // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details".
    // The "Author" built-in property is in the "Origin" group, and the others are in the "Description" group.
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.Description.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.Description.docx"));
    
    properties = doc->get_BuiltInDocumentProperties();
    
    ASSERT_EQ(properties->get_Author(), u"John Doe");
    ASSERT_EQ(properties->get_Category(), u"My category");
    ASSERT_EQ(properties->get_Comments(), System::String::Format(u"This is {0}'s document about {1}", properties->get_Author(), properties->get_Subject()));
    ASSERT_EQ(properties->get_Keywords(), u"Tag 1; Tag 2; Tag 3");
    ASSERT_EQ(properties->get_Subject(), u"My subject");
    ASSERT_EQ(properties->get_Title(), u"John's Document");
    ASSERT_EQ(doc->GetText().Trim(), System::String(u"Author:\t\u0013 AUTHOR \u0014John Doe\u0015\r") + u"Doc title:\t\u0013 TITLE \u0014John's Document\u0015\r" + u"Subject:\t\u0013 SUBJECT \u0014My subject\u0015\r" + u"Comments:\t\"\u0013 COMMENTS \u0014This is John Doe's document about My subject\u0015\"");
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, Description)
{
    s_instance->Description();
}

} // namespace gtest_test

void ExDocumentProperties::Origin()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.Company
    //ExFor:BuiltInDocumentProperties.CreatedTime
    //ExFor:BuiltInDocumentProperties.LastPrinted
    //ExFor:BuiltInDocumentProperties.LastSavedBy
    //ExFor:BuiltInDocumentProperties.LastSavedTime
    //ExFor:BuiltInDocumentProperties.Manager
    //ExFor:BuiltInDocumentProperties.NameOfApplication
    //ExFor:BuiltInDocumentProperties.RevisionNumber
    //ExFor:BuiltInDocumentProperties.Template
    //ExFor:BuiltInDocumentProperties.TotalEditingTime
    //ExFor:BuiltInDocumentProperties.Version
    //ExSummary:Shows how to work with document properties in the "Origin" category.
    // Open a document that we have created and edited using Microsoft Word.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Properties.docx"));
    System::SharedPtr<Aspose::Words::Properties::BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
    
    // The following built-in properties contain information regarding the creation and editing of this document.
    // We can right-click this document in Windows Explorer and find
    // these properties via "Properties" -> "Details" -> "Origin" category.
    // Fields such as PRINTDATE and EDITTIME can display these values in the document body.
    std::cout << System::String::Format(u"Created using {0}, on {1}", properties->get_NameOfApplication(), properties->get_CreatedTime()) << std::endl;
    std::cout << System::String::Format(u"Minutes spent editing: {0}", properties->get_TotalEditingTime()) << std::endl;
    std::cout << System::String::Format(u"Date/time last printed: {0}", properties->get_LastPrinted()) << std::endl;
    std::cout << System::String::Format(u"Template document: {0}", properties->get_Template()) << std::endl;
    
    // We can also change the values of built-in properties.
    properties->set_Company(u"Doe Ltd.");
    properties->set_Manager(u"Jane Doe");
    properties->set_Version(5);
    System::WithLambda::setter_post_increment_wrap(GETTER_SETTER_LAMBDA_ARGS(properties, RevisionNumber));
    
    // Microsoft Word updates the following properties automatically when we save the document.
    // To use these properties with Aspose.Words, we will need to set values for them manually.
    properties->set_LastSavedBy(u"John Doe");
    properties->set_LastSavedTime(System::DateTime::get_Now());
    
    // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details" -> "Origin".
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.Origin.docx");
    //ExEnd
    
    properties = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.Origin.docx"))->get_BuiltInDocumentProperties();
    
    ASSERT_EQ(properties->get_Company(), u"Doe Ltd.");
    ASSERT_EQ(properties->get_CreatedTime(), System::DateTime(2006, 4, 25, 10, 10, 0));
    ASSERT_EQ(properties->get_LastPrinted(), System::DateTime(2019, 4, 21, 10, 0, 0));
    ASSERT_EQ(properties->get_LastSavedBy(), u"John Doe");
    TestUtil::VerifyDate(System::DateTime::get_Now(), properties->get_LastSavedTime(), System::TimeSpan::FromSeconds(5));
    ASSERT_EQ(properties->get_Manager(), u"Jane Doe");
    ASSERT_EQ(properties->get_NameOfApplication(), u"Microsoft Office Word");
    ASSERT_EQ(properties->get_RevisionNumber(), 12);
    ASSERT_EQ(properties->get_Template(), u"Normal");
    ASSERT_EQ(properties->get_TotalEditingTime(), 8);
    ASSERT_EQ(properties->get_Version(), 786432);
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, Origin)
{
    s_instance->Origin();
}

} // namespace gtest_test

void ExDocumentProperties::Content()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Paragraphs.docx"));
    System::SharedPtr<Aspose::Words::Properties::BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
    
    // By using built in properties,
    // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
    // These properties are accessed by right clicking the file in Windows Explorer and navigating to Properties > Details > Content
    // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
    // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
    // Page count: The PageCount property shows the page count in real time and its value can be assigned to the Pages property
    
    // The "Pages" property stores the page count of the document.
    ASSERT_EQ(properties->get_Pages(), 6);
    
    // The "Words", "Characters", and "CharactersWithSpaces" built-in properties also display various document statistics,
    // but we need to call the "UpdateWordCount" method on the whole document before we can expect them to contain accurate values.
    ASSERT_EQ(properties->get_Words(), 1054);
    //ExSkip
    ASSERT_EQ(properties->get_Characters(), 6009);
    //ExSkip
    ASSERT_EQ(properties->get_CharactersWithSpaces(), 7049);
    //ExSkip
    doc->UpdateWordCount();
    
    ASSERT_EQ(properties->get_Words(), 1035);
    ASSERT_EQ(properties->get_Characters(), 6026);
    ASSERT_EQ(properties->get_CharactersWithSpaces(), 7041);
    
    // Count the number of lines in the document, and then assign the result to the "Lines" built-in property.
    auto lineCounter = System::MakeObject<Aspose::Words::ApiExamples::ExDocumentProperties::LineCounter>(doc);
    properties->set_Lines(lineCounter->GetLineCount());
    
    ASSERT_EQ(properties->get_Lines(), 142);
    
    // Assign the number of Paragraph nodes in the document to the "Paragraphs" built-in property.
    properties->set_Paragraphs(doc->GetChildNodes(NodeType::Paragraph, true)->get_Count());
    ASSERT_EQ(properties->get_Paragraphs(), 29);
    
    // Get an estimate of the file size of our document via the "Bytes" built-in property.
    ASSERT_EQ(properties->get_Bytes(), 20310);
    
    // Set a different template for our document, and then update the "Template" built-in property manually to reflect this change.
    doc->set_AttachedTemplate(get_MyDir() + u"Business brochure.dotx");
    
    ASSERT_EQ(properties->get_Template(), u"Normal");
    
    properties->set_Template(doc->get_AttachedTemplate());
    
    // "ContentStatus" is a descriptive built-in property.
    properties->set_ContentStatus(u"Draft");
    
    // Upon saving, the "ContentType" built-in property will contain the MIME type of the output save format.
    ASSERT_EQ(properties->get_ContentType(), System::String::Empty);
    
    // If the document contains links, and they are all up to date, we can set the "LinksUpToDate" property to "true".
    ASSERT_FALSE(properties->get_LinksUpToDate());
    
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.Content.docx");
    TestContent(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.Content.docx")));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, Content)
{
    s_instance->Content();
}

} // namespace gtest_test

void ExDocumentProperties::Thumbnail()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.Thumbnail
    //ExFor:DocumentProperty.ToByteArray
    //ExSummary:Shows how to add a thumbnail to a document that we save as an Epub.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // If we save a document, whose "Thumbnail" property contains image data that we added, as an Epub,
    // a reader that opens that document may display the image before the first page.
    System::SharedPtr<Aspose::Words::Properties::BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
    
    System::ArrayPtr<uint8_t> thumbnailBytes = System::IO::File::ReadAllBytes(get_ImageDir() + u"Logo.jpg");
    properties->set_Thumbnail(thumbnailBytes);
    
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.Thumbnail.epub");
    
    // We can extract a document's thumbnail image and save it to the local file system.
    System::SharedPtr<Aspose::Words::Properties::DocumentProperty> thumbnail = doc->get_BuiltInDocumentProperties()->idx_get(u"Thumbnail");
    System::IO::File::WriteAllBytes(get_ArtifactsDir() + u"DocumentProperties.Thumbnail.gif", thumbnail->ToByteArray());
    //ExEnd
    
    TestUtil::VerifyImage(400, 400, get_ArtifactsDir() + u"DocumentProperties.Thumbnail.gif");
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, Thumbnail)
{
    s_instance->Thumbnail();
}

} // namespace gtest_test

void ExDocumentProperties::HyperlinkBase()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.HyperlinkBase
    //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a relative hyperlink to a document in the local file system named "Document.docx".
    // Clicking on the link in Microsoft Word will open the designated document, if it is available.
    builder->InsertHyperlink(u"Relative hyperlink", u"Document.docx", false);
    
    // This link is relative. If there is no "Document.docx" in the same folder
    // as the document that contains this link, the link will be broken.
    ASSERT_FALSE(System::IO::File::Exists(get_ArtifactsDir() + u"Document.docx"));
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.HyperlinkBase.BrokenLink.docx");
    
    // The document we are trying to link to is in a different directory to the one we are planning to save the document in.
    // We could fix links like this by putting an absolute filename in each one.
    // Alternatively, we could provide a base link that every hyperlink with a relative filename
    // will prepend to its link when we click on it.
    System::SharedPtr<Aspose::Words::Properties::BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
    properties->set_HyperlinkBase(get_MyDir());
    
    ASSERT_TRUE(System::IO::File::Exists(properties->get_HyperlinkBase() + (System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address()));
    
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.HyperlinkBase.WorkingLink.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.HyperlinkBase.BrokenLink.docx"));
    properties = doc->get_BuiltInDocumentProperties();
    
    ASSERT_EQ(properties->get_HyperlinkBase(), System::String::Empty);
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.HyperlinkBase.WorkingLink.docx"));
    properties = doc->get_BuiltInDocumentProperties();
    
    ASSERT_EQ(properties->get_HyperlinkBase(), get_MyDir());
    ASSERT_TRUE(System::IO::File::Exists(properties->get_HyperlinkBase() + (System::ExplicitCast<Aspose::Words::Fields::FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address()));
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, HyperlinkBase)
{
    s_instance->HyperlinkBase();
}

} // namespace gtest_test

void ExDocumentProperties::HeadingPairs()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.HeadingPairs
    //ExFor:BuiltInDocumentProperties.TitlesOfParts
    //ExSummary:Shows the relationship between "HeadingPairs" and "TitlesOfParts" properties.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Heading pairs and titles of parts.docx"));
    
    // We can find the combined values of these collections via
    // "File" -> "Properties" -> "Advanced Properties" -> "Contents" tab.
    // The HeadingPairs property is a collection of <string, int> pairs that
    // determines how many document parts a heading spans across.
    System::ArrayPtr<System::SharedPtr<System::Object>> headingPairs = doc->get_BuiltInDocumentProperties()->get_HeadingPairs();
    
    // The TitlesOfParts property contains the names of parts that belong to the above headings.
    System::ArrayPtr<System::String> titlesOfParts = doc->get_BuiltInDocumentProperties()->get_TitlesOfParts();
    
    int32_t headingPairsIndex = 0;
    int32_t titlesOfPartsIndex = 0;
    while (headingPairsIndex < headingPairs->get_Length())
    {
        std::cout << System::String::Format(u"Parts for {0}:", headingPairs[headingPairsIndex++]) << std::endl;
        int32_t partsCount = System::Convert::ToInt32(headingPairs[headingPairsIndex++]);
        
        for (int32_t i = 0; i < partsCount; i++)
        {
            std::cout << System::String::Format(u"\t\"{0}\"", titlesOfParts[titlesOfPartsIndex++]) << std::endl;
        }
    }
    //ExEnd
    
    // There are 6 array elements designating 3 heading/part count pairs
    ASSERT_EQ(headingPairs->get_Length(), 6);
    ASSERT_EQ(System::ObjectExt::ToString(headingPairs[0]), u"Title");
    ASSERT_EQ(System::ObjectExt::ToString(headingPairs[1]), u"1");
    ASSERT_EQ(System::ObjectExt::ToString(headingPairs[2]), u"Heading 1");
    ASSERT_EQ(System::ObjectExt::ToString(headingPairs[3]), u"5");
    ASSERT_EQ(System::ObjectExt::ToString(headingPairs[4]), u"Heading 2");
    ASSERT_EQ(System::ObjectExt::ToString(headingPairs[5]), u"2");
    
    ASSERT_EQ(titlesOfParts->get_Length(), 8);
    // "Title"
    ASSERT_EQ(titlesOfParts[0], u"");
    // "Heading 1"
    ASSERT_EQ(titlesOfParts[1], u"Part1");
    ASSERT_EQ(titlesOfParts[2], u"Part2");
    ASSERT_EQ(titlesOfParts[3], u"Part3");
    ASSERT_EQ(titlesOfParts[4], u"Part4");
    ASSERT_EQ(titlesOfParts[5], u"Part5");
    // "Heading 2"
    ASSERT_EQ(titlesOfParts[6], u"Part6");
    ASSERT_EQ(titlesOfParts[7], u"Part7");
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, HeadingPairs)
{
    s_instance->HeadingPairs();
}

} // namespace gtest_test

void ExDocumentProperties::Security()
{
    //ExStart
    //ExFor:BuiltInDocumentProperties.Security
    //ExFor:DocumentSecurity
    //ExSummary:Shows how to use document properties to display the security level of a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_EQ(doc->get_BuiltInDocumentProperties()->get_Security(), DocumentSecurity::None);
    
    // If we configure a document to be read-only, it will display this status using the "Security" built-in property.
    doc->get_WriteProtection()->set_ReadOnlyRecommended(true);
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.Security.ReadOnlyRecommended.docx");
    
    ASSERT_EQ(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.Security.ReadOnlyRecommended.docx"))->get_BuiltInDocumentProperties()->get_Security(), DocumentSecurity::ReadOnlyRecommended);
    
    // Write-protect a document, and then verify its security level.
    doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_FALSE(doc->get_WriteProtection()->get_IsWriteProtected());
    
    doc->get_WriteProtection()->SetPassword(u"MyPassword");
    
    ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));
    ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());
    
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.Security.ReadOnlyEnforced.docx");
    
    ASSERT_EQ(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.Security.ReadOnlyEnforced.docx"))->get_BuiltInDocumentProperties()->get_Security(), DocumentSecurity::ReadOnlyEnforced);
    
    // "Security" is a descriptive property. We can edit its value manually.
    doc = System::MakeObject<Aspose::Words::Document>();
    
    doc->Protect(ProtectionType::AllowOnlyComments, u"MyPassword");
    doc->get_BuiltInDocumentProperties()->set_Security(DocumentSecurity::ReadOnlyExceptAnnotations);
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.Security.ReadOnlyExceptAnnotations.docx");
    
    ASSERT_EQ(System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.Security.ReadOnlyExceptAnnotations.docx"))->get_BuiltInDocumentProperties()->get_Security(), DocumentSecurity::ReadOnlyExceptAnnotations);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, Security)
{
    s_instance->Security();
}

} // namespace gtest_test

void ExDocumentProperties::CustomNamedAccess()
{
    //ExStart
    //ExFor:DocumentPropertyCollection.Item(String)
    //ExFor:CustomDocumentProperties.Add(String,DateTime)
    //ExFor:DocumentProperty.ToDateTime
    //ExSummary:Shows how to create a custom document property which contains a date and time.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    doc->get_CustomDocumentProperties()->Add(u"AuthorizationDate", System::DateTime::get_Now());
    System::DateTime authorizationDate = doc->get_CustomDocumentProperties()->idx_get(u"AuthorizationDate")->ToDateTime();
    std::cout << System::String::Format(u"Document authorized on {0}", authorizationDate) << std::endl;
    //ExEnd
    
    TestUtil::VerifyDate(System::DateTime::get_Now(), DocumentHelper::SaveOpen(doc)->get_CustomDocumentProperties()->idx_get(u"AuthorizationDate")->ToDateTime(), System::TimeSpan::FromSeconds(1));
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, CustomNamedAccess)
{
    s_instance->CustomNamedAccess();
}

} // namespace gtest_test

void ExDocumentProperties::LinkCustomDocumentPropertiesToBookmark()
{
    //ExStart
    //ExFor:CustomDocumentProperties.AddLinkToContent(String, String)
    //ExFor:DocumentProperty.IsLinkToContent
    //ExFor:DocumentProperty.LinkSource
    //ExSummary:Shows how to link a custom document property to a bookmark.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->StartBookmark(u"MyBookmark");
    builder->Write(u"Hello world!");
    builder->EndBookmark(u"MyBookmark");
    
    // Link a new custom property to a bookmark. The value of this property
    // will be the contents of the bookmark that it references in the "LinkSource" member.
    System::SharedPtr<Aspose::Words::Properties::CustomDocumentProperties> customProperties = doc->get_CustomDocumentProperties();
    System::SharedPtr<Aspose::Words::Properties::DocumentProperty> customProperty = customProperties->AddLinkToContent(u"Bookmark", u"MyBookmark");
    
    ASPOSE_ASSERT_EQ(customProperty->get_IsLinkToContent(), true);
    ASSERT_EQ(customProperty->get_LinkSource(), u"MyBookmark");
    ASPOSE_ASSERT_EQ(customProperty->get_Value(), u"Hello world!");
    
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(System::String(get_ArtifactsDir() + u"DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx"));
    customProperty = doc->get_CustomDocumentProperties()->idx_get(u"Bookmark");
    
    ASPOSE_ASSERT_EQ(customProperty->get_IsLinkToContent(), true);
    ASSERT_EQ(customProperty->get_LinkSource(), u"MyBookmark");
    ASPOSE_ASSERT_EQ(customProperty->get_Value(), u"Hello world!");
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, LinkCustomDocumentPropertiesToBookmark)
{
    s_instance->LinkCustomDocumentPropertiesToBookmark();
}

} // namespace gtest_test

void ExDocumentProperties::DocumentPropertyCollection()
{
    //ExStart
    //ExFor:CustomDocumentProperties.Add(String,String)
    //ExFor:CustomDocumentProperties.Add(String,Boolean)
    //ExFor:CustomDocumentProperties.Add(String,int)
    //ExFor:CustomDocumentProperties.Add(String,DateTime)
    //ExFor:CustomDocumentProperties.Add(String,Double)
    //ExFor:DocumentProperty.Type
    //ExFor:DocumentPropertyCollection
    //ExFor:DocumentPropertyCollection.Clear
    //ExFor:DocumentPropertyCollection.Contains(String)
    //ExFor:DocumentPropertyCollection.GetEnumerator
    //ExFor:DocumentPropertyCollection.IndexOf(String)
    //ExFor:DocumentPropertyCollection.RemoveAt(Int32)
    //ExFor:DocumentPropertyCollection.Remove
    //ExFor:PropertyType
    //ExSummary:Shows how to work with a document's custom properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Properties::CustomDocumentProperties> properties = doc->get_CustomDocumentProperties();
    
    ASSERT_EQ(properties->get_Count(), 0);
    
    // Custom document properties are key-value pairs that we can add to the document.
    properties->Add(u"Authorized", true);
    properties->Add(u"Authorized By", System::String(u"John Doe"));
    properties->Add(u"Authorized Date", System::DateTime::get_Today());
    properties->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
    properties->Add(u"Authorized Amount", 123.45);
    
    // The collection sorts the custom properties in alphabetic order.
    ASSERT_EQ(properties->IndexOf(u"Authorized Amount"), 1);
    ASSERT_EQ(properties->get_Count(), 5);
    
    // Print every custom property in the document.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Properties::DocumentProperty>>> enumerator = properties->GetEnumerator();
        while (enumerator->MoveNext())
        {
            std::cout << System::String::Format(u"Name: \"{0}\"\n\tType: \"{1}\"\n\tValue: \"{2}\"", enumerator->get_Current()->get_Name(), enumerator->get_Current()->get_Type(), enumerator->get_Current()->get_Value()) << std::endl;
        }
    }
    
    // Display the value of a custom property using a DOCPROPERTY field.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    auto field = System::ExplicitCast<Aspose::Words::Fields::FieldDocProperty>(builder->InsertField(u" DOCPROPERTY \"Authorized By\""));
    field->Update();
    
    ASSERT_EQ(field->get_Result(), u"John Doe");
    
    // We can find these custom properties in Microsoft Word via "File" -> "Properties" > "Advanced Properties" > "Custom".
    doc->Save(get_ArtifactsDir() + u"DocumentProperties.DocumentPropertyCollection.docx");
    
    // Below are three ways or removing custom properties from a document.
    // 1 -  Remove by index:
    properties->RemoveAt(1);
    
    ASSERT_FALSE(properties->Contains(u"Authorized Amount"));
    ASSERT_EQ(properties->get_Count(), 4);
    
    // 2 -  Remove by name:
    properties->Remove(u"Authorized Revision");
    
    ASSERT_FALSE(properties->Contains(u"Authorized Revision"));
    ASSERT_EQ(properties->get_Count(), 3);
    
    // 3 -  Empty the entire collection at once:
    properties->Clear();
    
    ASSERT_EQ(properties->get_Count(), 0);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, DocumentPropertyCollection)
{
    s_instance->DocumentPropertyCollection();
}

} // namespace gtest_test

void ExDocumentProperties::PropertyTypes()
{
    //ExStart
    //ExFor:DocumentProperty.ToBool
    //ExFor:DocumentProperty.ToInt
    //ExFor:DocumentProperty.ToDouble
    //ExFor:DocumentProperty.ToString
    //ExFor:DocumentProperty.ToDateTime
    //ExSummary:Shows various type conversion methods of custom document properties.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Properties::CustomDocumentProperties> properties = doc->get_CustomDocumentProperties();
    
    System::DateTime authDate = System::DateTime::get_Today();
    properties->Add(u"Authorized", true);
    properties->Add(u"Authorized By", System::String(u"John Doe"));
    properties->Add(u"Authorized Date", authDate);
    properties->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
    properties->Add(u"Authorized Amount", 123.45);
    
    ASPOSE_ASSERT_EQ(properties->idx_get(u"Authorized")->ToBool(), true);
    ASSERT_EQ(System::ObjectExt::ToString(properties->idx_get(u"Authorized By")), u"John Doe");
    ASSERT_EQ(properties->idx_get(u"Authorized Date")->ToDateTime(), authDate);
    ASSERT_EQ(properties->idx_get(u"Authorized Revision")->ToInt(), 1);
    ASPOSE_ASSERT_EQ(properties->idx_get(u"Authorized Amount")->ToDouble(), 123.45);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, PropertyTypes)
{
    s_instance->PropertyTypes();
}

} // namespace gtest_test

void ExDocumentProperties::ExtendedProperties()
{
    //ExStart:ExtendedProperties
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:BuiltInDocumentProperties.ScaleCrop
    //ExFor:BuiltInDocumentProperties.SharedDocument
    //ExFor:BuiltInDocumentProperties.HyperlinksChanged
    //ExSummary:Shows how to get extended properties.
    auto doc = System::MakeObject<Aspose::Words::Document>(System::String(get_MyDir() + u"Extended properties.docx"));
    ASSERT_TRUE(doc->get_BuiltInDocumentProperties()->get_ScaleCrop());
    ASSERT_TRUE(doc->get_BuiltInDocumentProperties()->get_SharedDocument());
    ASSERT_TRUE(doc->get_BuiltInDocumentProperties()->get_HyperlinksChanged());
    //ExEnd:ExtendedProperties
}

namespace gtest_test
{

TEST_F(ExDocumentProperties, ExtendedProperties)
{
    s_instance->ExtendedProperties();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
