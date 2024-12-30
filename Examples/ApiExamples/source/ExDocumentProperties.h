#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldDocProperty.h>
#include <Aspose.Words.Cpp/Fields/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Layout/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Properties/BuiltInDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/CustomDocumentProperties.h>
#include <Aspose.Words.Cpp/Properties/DocumentProperty.h>
#include <Aspose.Words.Cpp/Properties/DocumentSecurity.h>
#include <Aspose.Words.Cpp/Properties/PropertyType.h>
#include <Aspose.Words.Cpp/ProtectionType.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Settings/WriteProtection.h>
#include <system/array.h>
#include <system/collections/icollection.h>
#include <system/collections/ienumerator.h>
#include <system/convert.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enum.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/io/file_mode.h>
#include <system/io/file_stream.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/timespan.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Layout;
using namespace Aspose::Words::Properties;

namespace ApiExamples {

class ExDocumentProperties : public ApiExampleBase
{
public:
    void BuiltIn()
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
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        // The "Document" object contains some of its metadata in its members.
        std::cout << "Document filename:\n\t \"" << doc->get_OriginalFileName() << "\"" << std::endl;

        // The document also stores metadata in its built-in properties.
        // Each built-in property is a member of the document's "BuiltInDocumentProperties" object.
        std::cout << "Built-in Properties:" << std::endl;
        for (const auto& docProperty : System::IterateOver(doc->get_BuiltInDocumentProperties()))
        {
            std::cout << docProperty->get_Name() << std::endl;
            std::cout << String::Format(u"\tType:\t{0}", docProperty->get_Type()) << std::endl;

            // Some properties may store multiple values.
            if (System::ObjectExt::Is<System::Collections::Generic::ICollection<SharedPtr<System::Object>>>(docProperty->get_Value()))
            {
                for (const auto& value : System::IterateOver(
                         System::AsCast<System::Collections::Generic::ICollection<SharedPtr<System::Object>>>(docProperty->get_Value())))
                {
                    std::cout << "\tValue:\t\"" << value << "\"" << std::endl;
                }
            }
            else
            {
                std::cout << "\tValue:\t\"" << docProperty->get_Value() << "\"" << std::endl;
            }
        }
        //ExEnd

        ASSERT_EQ(31, doc->get_BuiltInDocumentProperties()->get_Count());
    }

    void Custom()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Item(String)
        //ExFor:CustomDocumentProperties
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentPropertyCollection.Count
        //ExFor:DocumentPropertyCollection.Item(int)
        //ExSummary:Shows how to work with custom document properties.
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");

        // Every document contains a collection of custom properties, which, like the built-in properties, are key-value pairs.
        // The document has a fixed list of built-in properties. The user creates all of the custom properties.
        ASSERT_EQ(u"Value of custom document property", System::ObjectExt::ToString(doc->get_CustomDocumentProperties()->idx_get(u"CustomProperty")));

        doc->get_CustomDocumentProperties()->Add(u"CustomProperty2", String(u"Value of custom document property #2"));

        std::cout << "Custom Properties:" << std::endl;
        for (const auto& customDocumentProperty : System::IterateOver(doc->get_CustomDocumentProperties()))
        {
            std::cout << customDocumentProperty->get_Name() << std::endl;
            std::cout << String::Format(u"\tType:\t{0}", customDocumentProperty->get_Type()) << std::endl;
            std::cout << "\tValue:\t\"" << customDocumentProperty->get_Value() << "\"" << std::endl;
        }
        //ExEnd

        ASSERT_EQ(2, doc->get_CustomDocumentProperties()->get_Count());
    }

    void Description()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Author
        //ExFor:BuiltInDocumentProperties.Category
        //ExFor:BuiltInDocumentProperties.Comments
        //ExFor:BuiltInDocumentProperties.Keywords
        //ExFor:BuiltInDocumentProperties.Subject
        //ExFor:BuiltInDocumentProperties.Title
        //ExSummary:Shows how to work with built-in document properties in the "Description" category.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

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
        properties->set_Comments(String::Format(u"This is {0}'s document about {1}", properties->get_Author(), properties->get_Subject()));
        builder->Write(u"\nComments:\t\"");
        builder->InsertField(FieldType::FieldComments, true);
        builder->Write(u"\"");

        // The "Category" built-in property does not have a field that can display its value.
        properties->set_Category(u"My category");

        // We can set multiple keywords for a document by separating the string value of the "Keywords" property with semicolons.
        properties->set_Keywords(u"Tag 1; Tag 2; Tag 3");

        // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details".
        // The "Author" built-in property is in the "Origin" group, and the others are in the "Description" group.
        doc->Save(ArtifactsDir + u"DocumentProperties.Description.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentProperties.Description.docx");

        properties = doc->get_BuiltInDocumentProperties();

        ASSERT_EQ(u"John Doe", properties->get_Author());
        ASSERT_EQ(u"My category", properties->get_Category());
        ASSERT_EQ(String::Format(u"This is {0}'s document about {1}", properties->get_Author(), properties->get_Subject()), properties->get_Comments());
        ASSERT_EQ(u"Tag 1; Tag 2; Tag 3", properties->get_Keywords());
        ASSERT_EQ(u"My subject", properties->get_Subject());
        ASSERT_EQ(u"John's Document", properties->get_Title());
        ASSERT_EQ(String(u"Author:\t\u0013 AUTHOR \u0014John Doe\u0015\r") + u"Doc title:\t\u0013 TITLE \u0014John's Document\u0015\r" +
                      u"Subject:\t\u0013 SUBJECT \u0014My subject\u0015\r" +
                      u"Comments:\t\"\u0013 COMMENTS \u0014This is John Doe's document about My subject\u0015\"",
                  doc->GetText().Trim());
    }

    void Origin()
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
        auto doc = MakeObject<Document>(MyDir + u"Properties.docx");
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        // The following built-in properties contain information regarding the creation and editing of this document.
        // We can right-click this document in Windows Explorer and find
        // these properties via "Properties" -> "Details" -> "Origin" category.
        // Fields such as PRINTDATE and EDITTIME can display these values in the document body.
        std::cout << "Created using " << properties->get_NameOfApplication() << ", on " << properties->get_CreatedTime() << std::endl;
        std::cout << "Minutes spent editing: " << properties->get_TotalEditingTime() << std::endl;
        std::cout << "Date/time last printed: " << properties->get_LastPrinted() << std::endl;
        std::cout << "Template document: " << properties->get_Template() << std::endl;

        // We can also change the values of built-in properties.
        properties->set_Company(u"Doe Ltd.");
        properties->set_Manager(u"Jane Doe");
        properties->set_Version(5);
        properties->set_RevisionNumber(properties->get_RevisionNumber() + 1);

        // Microsoft Word updates the following properties automatically when we save the document.
        // To use these properties with Aspose.Words, we will need to set values for them manually.
        properties->set_LastSavedBy(u"John Doe");
        properties->set_LastSavedTime(System::DateTime::get_Now());

        // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details" -> "Origin".
        doc->Save(ArtifactsDir + u"DocumentProperties.Origin.docx");
        //ExEnd

        properties = MakeObject<Document>(ArtifactsDir + u"DocumentProperties.Origin.docx")->get_BuiltInDocumentProperties();

        ASSERT_EQ(u"Doe Ltd.", properties->get_Company());
        ASSERT_EQ(System::DateTime(2006, 4, 25, 10, 10, 0), properties->get_CreatedTime());
        ASSERT_EQ(System::DateTime(2019, 4, 21, 10, 0, 0), properties->get_LastPrinted());
        ASSERT_EQ(u"John Doe", properties->get_LastSavedBy());
        TestUtil::VerifyDate(System::DateTime::get_Now(), properties->get_LastSavedTime(), System::TimeSpan::FromSeconds(5));
        ASSERT_EQ(u"Jane Doe", properties->get_Manager());
        ASSERT_EQ(u"Microsoft Office Word", properties->get_NameOfApplication());
        ASSERT_EQ(12, properties->get_RevisionNumber());
        ASSERT_EQ(u"Normal", properties->get_Template());
        ASSERT_EQ(8, properties->get_TotalEditingTime());
        ASSERT_EQ(786432, properties->get_Version());
    }

    //ExStart
    //ExFor:BuiltInDocumentProperties.Bytes
    //ExFor:BuiltInDocumentProperties.Characters
    //ExFor:BuiltInDocumentProperties.CharactersWithSpaces
    //ExFor:BuiltInDocumentProperties.ContentStatus
    //ExFor:BuiltInDocumentProperties.ContentType
    //ExFor:BuiltInDocumentProperties.Lines
    //ExFor:BuiltInDocumentProperties.LinksUpToDate
    //ExFor:BuiltInDocumentProperties.Pages
    //ExFor:BuiltInDocumentProperties.Paragraphs
    //ExFor:BuiltInDocumentProperties.Words
    //ExSummary:Shows how to work with document properties in the "Content" category.
    void Content()
    {
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        // By using built in properties,
        // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
        // These properties are accessed by right clicking the file in Windows Explorer and navigating to Properties > Details > Content
        // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
        // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
        // Page count: The PageCount property shows the page count in real time and its value can be assigned to the Pages property

        // The "Pages" property stores the page count of the document.
        ASSERT_EQ(6, properties->get_Pages());

        // The "Words", "Characters", and "CharactersWithSpaces" built-in properties also display various document statistics,
        // but we need to call the "UpdateWordCount" method on the whole document before we can expect them to contain accurate values.
        ASSERT_EQ(1054, properties->get_Words());
        //ExSkip
        ASSERT_EQ(6009, properties->get_Characters());
        //ExSkip
        ASSERT_EQ(7049, properties->get_CharactersWithSpaces());
        //ExSkip
        doc->UpdateWordCount();

        ASSERT_EQ(1035, properties->get_Words());
        ASSERT_EQ(6026, properties->get_Characters());
        ASSERT_EQ(7041, properties->get_CharactersWithSpaces());

        // Count the number of lines in the document, and then assign the result to the "Lines" built-in property.
        auto lineCounter = MakeObject<ExDocumentProperties::LineCounter>(doc);
        properties->set_Lines(lineCounter->GetLineCount());

        ASSERT_EQ(142, properties->get_Lines());

        // Assign the number of Paragraph nodes in the document to the "Paragraphs" built-in property.
        properties->set_Paragraphs(doc->GetChildNodes(NodeType::Paragraph, true)->get_Count());
        ASSERT_EQ(29, properties->get_Paragraphs());

        // Get an estimate of the file size of our document via the "Bytes" built-in property.
        ASSERT_EQ(20310, properties->get_Bytes());

        // Set a different template for our document, and then update the "Template" built-in property manually to reflect this change.
        doc->set_AttachedTemplate(MyDir + u"Business brochure.dotx");

        ASSERT_EQ(u"Normal", properties->get_Template());

        properties->set_Template(doc->get_AttachedTemplate());

        // "ContentStatus" is a descriptive built-in property.
        properties->set_ContentStatus(u"Draft");

        // Upon saving, the "ContentType" built-in property will contain the MIME type of the output save format.
        ASSERT_EQ(String::Empty, properties->get_ContentType());

        // If the document contains links, and they are all up to date, we can set the "LinksUpToDate" property to "true".
        ASSERT_FALSE(properties->get_LinksUpToDate());

        doc->Save(ArtifactsDir + u"DocumentProperties.Content.docx");
        TestContent(MakeObject<Document>(ArtifactsDir + u"DocumentProperties.Content.docx"));
        //ExSkip
    }

    /// <summary>
    /// Counts the lines in a document.
    /// Traverses the document's layout entities tree upon construction,
    /// counting entities of the "Line" type that also contain real text.
    /// </summary>
    class LineCounter : public System::Object
    {
    public:
        LineCounter(SharedPtr<Document> doc) : mLineCount(0), mScanningLineForRealText(false)
        {
            mLayoutEnumerator = MakeObject<LayoutEnumerator>(doc);

            CountLines();
        }

        int GetLineCount()
        {
            return mLineCount;
        }

    private:
        SharedPtr<LayoutEnumerator> mLayoutEnumerator;
        int mLineCount;
        bool mScanningLineForRealText;

        void CountLines()
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
    };
    //ExEnd

    void TestContent(SharedPtr<Document> doc)
    {
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        ASSERT_EQ(6, properties->get_Pages());

        ASSERT_EQ(1035, properties->get_Words());
        ASSERT_EQ(6026, properties->get_Characters());
        ASSERT_EQ(7041, properties->get_CharactersWithSpaces());
        ASSERT_EQ(142, properties->get_Lines());
        ASSERT_EQ(29, properties->get_Paragraphs());
        ASSERT_NEAR(15500, properties->get_Bytes(), 200);
        ASSERT_EQ(MyDir.Replace(u"\\\\", u"\\") + u"Business brochure.dotx", properties->get_Template());
        ASSERT_EQ(u"Draft", properties->get_ContentStatus());
        ASSERT_EQ(String::Empty, properties->get_ContentType());
        ASSERT_FALSE(properties->get_LinksUpToDate());
    }

    void Thumbnail()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Thumbnail
        //ExFor:DocumentProperty.ToByteArray
        //ExSummary:Shows how to add a thumbnail to a document that we save as an Epub.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // If we save a document, whose "Thumbnail" property contains image data that we added, as an Epub,
        // a reader that opens that document may display the image before the first page.
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();

        ArrayPtr<uint8_t> thumbnailBytes = System::IO::File::ReadAllBytes(ImageDir + u"Logo.jpg");
        properties->set_Thumbnail(thumbnailBytes);

        doc->Save(ArtifactsDir + u"DocumentProperties.Thumbnail.epub");

        // We can extract a document's thumbnail image and save it to the local file system.
        SharedPtr<DocumentProperty> thumbnail = doc->get_BuiltInDocumentProperties()->idx_get(u"Thumbnail");
        System::IO::File::WriteAllBytes(ArtifactsDir + u"DocumentProperties.Thumbnail.gif", thumbnail->ToByteArray());
        //ExEnd

        {
            auto imgStream = MakeObject<System::IO::FileStream>(ArtifactsDir + u"DocumentProperties.Thumbnail.gif", System::IO::FileMode::Open);
            TestUtil::VerifyImage(400, 400, imgStream);
        }
    }

    void HyperlinkBase()
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.HyperlinkBase
        //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a relative hyperlink to a document in the local file system named "Document.docx".
        // Clicking on the link in Microsoft Word will open the designated document, if it is available.
        builder->InsertHyperlink(u"Relative hyperlink", u"Document.docx", false);

        // This link is relative. If there is no "Document.docx" in the same folder
        // as the document that contains this link, the link will be broken.
        ASSERT_FALSE(System::IO::File::Exists(ArtifactsDir + u"Document.docx"));
        doc->Save(ArtifactsDir + u"DocumentProperties.HyperlinkBase.BrokenLink.docx");

        // The document we are trying to link to is in a different directory to the one we are planning to save the document in.
        // We could fix links like this by putting an absolute filename in each one.
        // Alternatively, we could provide a base link that every hyperlink with a relative filename
        // will prepend to its link when we click on it.
        SharedPtr<BuiltInDocumentProperties> properties = doc->get_BuiltInDocumentProperties();
        properties->set_HyperlinkBase(MyDir);

        ASSERT_TRUE(System::IO::File::Exists(properties->get_HyperlinkBase() +
                                             (System::ExplicitCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address()));

        doc->Save(ArtifactsDir + u"DocumentProperties.HyperlinkBase.WorkingLink.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentProperties.HyperlinkBase.BrokenLink.docx");
        properties = doc->get_BuiltInDocumentProperties();

        ASSERT_EQ(String::Empty, properties->get_HyperlinkBase());

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentProperties.HyperlinkBase.WorkingLink.docx");
        properties = doc->get_BuiltInDocumentProperties();

        ASSERT_EQ(MyDir, properties->get_HyperlinkBase());
        ASSERT_TRUE(System::IO::File::Exists(properties->get_HyperlinkBase() +
                                             (System::ExplicitCast<FieldHyperlink>(doc->get_Range()->get_Fields()->idx_get(0)))->get_Address()));
    }

    void HeadingPairs()
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.HeadingPairs
        //ExFor:Properties.BuiltInDocumentProperties.TitlesOfParts
        //ExSummary:Shows the relationship between "HeadingPairs" and "TitlesOfParts" properties.
        auto doc = MakeObject<Document>(MyDir + u"Heading pairs and titles of parts.docx");

        // We can find the combined values of these collections via
        // "File" -> "Properties" -> "Advanced Properties" -> "Contents" tab.
        // The HeadingPairs property is a collection of <string, int> pairs that
        // determines how many document parts a heading spans across.
        ArrayPtr<SharedPtr<System::Object>> headingPairs = doc->get_BuiltInDocumentProperties()->get_HeadingPairs();

        // The TitlesOfParts property contains the names of parts that belong to the above headings.
        ArrayPtr<String> titlesOfParts = doc->get_BuiltInDocumentProperties()->get_TitlesOfParts();

        int headingPairsIndex = 0;
        int titlesOfPartsIndex = 0;
        while (headingPairsIndex < headingPairs->get_Length())
        {
            std::cout << "Parts for " << headingPairs[headingPairsIndex++] << ":" << std::endl;
            int partsCount = System::Convert::ToInt32(headingPairs[headingPairsIndex++]);

            for (int i = 0; i < partsCount; i++)
            {
                std::cout << "\t\"" << titlesOfParts[titlesOfPartsIndex++] << "\"" << std::endl;
            }
        }
        //ExEnd

        // There are 6 array elements designating 3 heading/part count pairs
        ASSERT_EQ(6, headingPairs->get_Length());
        ASSERT_EQ(u"Title", System::ObjectExt::ToString(headingPairs[0]));
        ASSERT_EQ(u"1", System::ObjectExt::ToString(headingPairs[1]));
        ASSERT_EQ(u"Heading 1", System::ObjectExt::ToString(headingPairs[2]));
        ASSERT_EQ(u"5", System::ObjectExt::ToString(headingPairs[3]));
        ASSERT_EQ(u"Heading 2", System::ObjectExt::ToString(headingPairs[4]));
        ASSERT_EQ(u"2", System::ObjectExt::ToString(headingPairs[5]));

        ASSERT_EQ(8, titlesOfParts->get_Length());
        // "Title"
        ASSERT_EQ(u"", titlesOfParts[0]);
        // "Heading 1"
        ASSERT_EQ(u"Part1", titlesOfParts[1]);
        ASSERT_EQ(u"Part2", titlesOfParts[2]);
        ASSERT_EQ(u"Part3", titlesOfParts[3]);
        ASSERT_EQ(u"Part4", titlesOfParts[4]);
        ASSERT_EQ(u"Part5", titlesOfParts[5]);
        // "Heading 2"
        ASSERT_EQ(u"Part6", titlesOfParts[6]);
        ASSERT_EQ(u"Part7", titlesOfParts[7]);
    }

    void Security()
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.Security
        //ExFor:Properties.DocumentSecurity
        //ExSummary:Shows how to use document properties to display the security level of a document.
        auto doc = MakeObject<Document>();

        ASSERT_EQ(DocumentSecurity::None, doc->get_BuiltInDocumentProperties()->get_Security());

        // If we configure a document to be read-only, it will display this status using the "Security" built-in property.
        doc->get_WriteProtection()->set_ReadOnlyRecommended(true);
        doc->Save(ArtifactsDir + u"DocumentProperties.Security.ReadOnlyRecommended.docx");

        ASSERT_EQ(
            DocumentSecurity::ReadOnlyRecommended,
            MakeObject<Document>(ArtifactsDir + u"DocumentProperties.Security.ReadOnlyRecommended.docx")->get_BuiltInDocumentProperties()->get_Security());

        // Write-protect a document, and then verify its security level.
        doc = MakeObject<Document>();

        ASSERT_FALSE(doc->get_WriteProtection()->get_IsWriteProtected());

        doc->get_WriteProtection()->SetPassword(u"MyPassword");

        ASSERT_TRUE(doc->get_WriteProtection()->ValidatePassword(u"MyPassword"));
        ASSERT_TRUE(doc->get_WriteProtection()->get_IsWriteProtected());

        doc->Save(ArtifactsDir + u"DocumentProperties.Security.ReadOnlyEnforced.docx");

        ASSERT_EQ(DocumentSecurity::ReadOnlyEnforced,
                  MakeObject<Document>(ArtifactsDir + u"DocumentProperties.Security.ReadOnlyEnforced.docx")->get_BuiltInDocumentProperties()->get_Security());

        // "Security" is a descriptive property. We can edit its value manually.
        doc = MakeObject<Document>();

        doc->Protect(ProtectionType::AllowOnlyComments, u"MyPassword");
        doc->get_BuiltInDocumentProperties()->set_Security(DocumentSecurity::ReadOnlyExceptAnnotations);
        doc->Save(ArtifactsDir + u"DocumentProperties.Security.ReadOnlyExceptAnnotations.docx");

        ASSERT_EQ(DocumentSecurity::ReadOnlyExceptAnnotations,
                  MakeObject<Document>(ArtifactsDir + u"DocumentProperties.Security.ReadOnlyExceptAnnotations.docx")
                      ->get_BuiltInDocumentProperties()
                      ->get_Security());
        //ExEnd
    }

    void CustomNamedAccess()
    {
        //ExStart
        //ExFor:DocumentPropertyCollection.Item(String)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Shows how to create a custom document property which contains a date and time.
        auto doc = MakeObject<Document>();

        doc->get_CustomDocumentProperties()->Add(u"AuthorizationDate", System::DateTime::get_Now());

        std::cout << "Document authorized on " << doc->get_CustomDocumentProperties()->idx_get(u"AuthorizationDate")->ToDateTime() << std::endl;
        //ExEnd

        TestUtil::VerifyDate(System::DateTime::get_Now(),
                             DocumentHelper::SaveOpen(doc)->get_CustomDocumentProperties()->idx_get(u"AuthorizationDate")->ToDateTime(),
                             System::TimeSpan::FromSeconds(1));
    }

    void LinkCustomDocumentPropertiesToBookmark()
    {
        //ExStart
        //ExFor:CustomDocumentProperties.AddLinkToContent(String, String)
        //ExFor:DocumentProperty.IsLinkToContent
        //ExFor:DocumentProperty.LinkSource
        //ExSummary:Shows how to link a custom document property to a bookmark.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartBookmark(u"MyBookmark");
        builder->Write(u"Hello world!");
        builder->EndBookmark(u"MyBookmark");

        // Link a new custom property to a bookmark. The value of this property
        // will be the contents of the bookmark that it references in the "LinkSource" member.
        SharedPtr<CustomDocumentProperties> customProperties = doc->get_CustomDocumentProperties();
        SharedPtr<DocumentProperty> customProperty = customProperties->AddLinkToContent(u"Bookmark", u"MyBookmark");

        ASPOSE_ASSERT_EQ(true, customProperty->get_IsLinkToContent());
        ASSERT_EQ(u"MyBookmark", customProperty->get_LinkSource());
        ASPOSE_ASSERT_EQ(u"Hello world!", customProperty->get_Value());

        doc->Save(ArtifactsDir + u"DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
        customProperty = doc->get_CustomDocumentProperties()->idx_get(u"Bookmark");

        ASPOSE_ASSERT_EQ(true, customProperty->get_IsLinkToContent());
        ASSERT_EQ(u"MyBookmark", customProperty->get_LinkSource());
        ASPOSE_ASSERT_EQ(u"Hello world!", customProperty->get_Value());
    }

    void DocumentPropertyCollection()
    {
        //ExStart
        //ExFor:CustomDocumentProperties.Add(String,String)
        //ExFor:CustomDocumentProperties.Add(String,Boolean)
        //ExFor:CustomDocumentProperties.Add(String,int)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:CustomDocumentProperties.Add(String,Double)
        //ExFor:DocumentProperty.Type
        //ExFor:Properties.DocumentPropertyCollection
        //ExFor:Properties.DocumentPropertyCollection.Clear
        //ExFor:Properties.DocumentPropertyCollection.Contains(System.String)
        //ExFor:Properties.DocumentPropertyCollection.GetEnumerator
        //ExFor:Properties.DocumentPropertyCollection.IndexOf(System.String)
        //ExFor:Properties.DocumentPropertyCollection.RemoveAt(System.Int32)
        //ExFor:Properties.DocumentPropertyCollection.Remove
        //ExFor:PropertyType
        //ExSummary:Shows how to work with a document's custom properties.
        auto doc = MakeObject<Document>();
        SharedPtr<CustomDocumentProperties> properties = doc->get_CustomDocumentProperties();

        ASSERT_EQ(0, properties->get_Count());

        // Custom document properties are key-value pairs that we can add to the document.
        properties->Add(u"Authorized", true);
        properties->Add(u"Authorized By", String(u"John Doe"));
        properties->Add(u"Authorized Date", System::DateTime::get_Today());
        properties->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
        properties->Add(u"Authorized Amount", 123.45);

        // The collection sorts the custom properties in alphabetic order.
        ASSERT_EQ(1, properties->IndexOf(u"Authorized Amount"));
        ASSERT_EQ(5, properties->get_Count());

        // Print every custom property in the document.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<DocumentProperty>>> enumerator = properties->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << String::Format(u"Name: \"{0}\"\n\tType: \"{1}\"\n\tValue: \"{2}\"", enumerator->get_Current()->get_Name(),
                                            enumerator->get_Current()->get_Type(), enumerator->get_Current()->get_Value())
                          << std::endl;
            }
        }

        // Display the value of a custom property using a DOCPROPERTY field.
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto field = System::ExplicitCast<FieldDocProperty>(builder->InsertField(u" DOCPROPERTY \"Authorized By\""));
        field->Update();

        ASSERT_EQ(u"John Doe", field->get_Result());

        // We can find these custom properties in Microsoft Word via "File" -> "Properties" > "Advanced Properties" > "Custom".
        doc->Save(ArtifactsDir + u"DocumentProperties.DocumentPropertyCollection.docx");

        // Below are three ways or removing custom properties from a document.
        // 1 -  Remove by index:
        properties->RemoveAt(1);

        ASSERT_FALSE(properties->Contains(u"Authorized Amount"));
        ASSERT_EQ(4, properties->get_Count());

        // 2 -  Remove by name:
        properties->Remove(u"Authorized Revision");

        ASSERT_FALSE(properties->Contains(u"Authorized Revision"));
        ASSERT_EQ(3, properties->get_Count());

        // 3 -  Empty the entire collection at once:
        properties->Clear();

        ASSERT_EQ(0, properties->get_Count());
        //ExEnd
    }

    void PropertyTypes()
    {
        //ExStart
        //ExFor:DocumentProperty.ToBool
        //ExFor:DocumentProperty.ToInt
        //ExFor:DocumentProperty.ToDouble
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Shows various type conversion methods of custom document properties.
        auto doc = MakeObject<Document>();
        SharedPtr<CustomDocumentProperties> properties = doc->get_CustomDocumentProperties();

        System::DateTime authDate = System::DateTime::get_Today();
        properties->Add(u"Authorized", true);
        properties->Add(u"Authorized By", String(u"John Doe"));
        properties->Add(u"Authorized Date", authDate);
        properties->Add(u"Authorized Revision", doc->get_BuiltInDocumentProperties()->get_RevisionNumber());
        properties->Add(u"Authorized Amount", 123.45);

        ASPOSE_ASSERT_EQ(true, properties->idx_get(u"Authorized")->ToBool());
        ASSERT_EQ(u"John Doe", System::ObjectExt::ToString(properties->idx_get(u"Authorized By")));
        ASSERT_EQ(authDate, properties->idx_get(u"Authorized Date")->ToDateTime());
        ASSERT_EQ(1, properties->idx_get(u"Authorized Revision")->ToInt());
        ASPOSE_ASSERT_EQ(123.45, properties->idx_get(u"Authorized Amount")->ToDouble());
        //ExEnd
    }
};

} // namespace ApiExamples
