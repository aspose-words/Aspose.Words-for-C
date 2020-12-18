#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/EditingLanguage.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldUpdateCultureSource.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DateAndTime/FieldTime.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/DocumentInformation/FieldFileName.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserAddress.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserInitials.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/UserInformation/FieldUserName.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormFieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUpdateCultureProvider.h>
#include <Aspose.Words.Cpp/Model/Fields/ToaCategories.h>
#include <Aspose.Words.Cpp/Model/Fields/UserInformation.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <system/array.h>
#include <system/exceptions.h>
#include <system/globalization/culture_info.h>
#include <system/globalization/date_time_format_info.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/threading/thread.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace ApiExamples {

class ExFieldOptions : public ApiExampleBase
{
public:
    void FieldOptionsCurrentUser()
    {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExFor:FieldOptions.CurrentUser
        //ExFor:UserInformation
        //ExFor:UserInformation.Name
        //ExFor:UserInformation.Initials
        //ExFor:UserInformation.Address
        //ExFor:UserInformation.DefaultUser
        //ExSummary:Shows how to set user details and display them with fields.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Set user information
        auto userInformation = MakeObject<UserInformation>();
        userInformation->set_Name(u"John Doe");
        userInformation->set_Initials(u"J. D.");
        userInformation->set_Address(u"123 Main Street");
        doc->get_FieldOptions()->set_CurrentUser(userInformation);

        // Insert fields that reference our user information
        ASSERT_EQ(userInformation->get_Name(), builder->InsertField(u" USERNAME ")->get_Result());
        ASSERT_EQ(userInformation->get_Initials(), builder->InsertField(u" USERINITIALS ")->get_Result());
        ASSERT_EQ(userInformation->get_Address(), builder->InsertField(u" USERADDRESS ")->get_Result());

        // The field options object also has a static default user value that fields from many documents can refer to
        UserInformation::get_DefaultUser()->set_Name(u"Default User");
        UserInformation::get_DefaultUser()->set_Initials(u"D. U.");
        UserInformation::get_DefaultUser()->set_Address(u"One Microsoft Way");
        doc->get_FieldOptions()->set_CurrentUser(UserInformation::get_DefaultUser());

        ASSERT_EQ(u"Default User", builder->InsertField(u" USERNAME ")->get_Result());
        ASSERT_EQ(u"D. U.", builder->InsertField(u" USERINITIALS ")->get_Result());
        ASSERT_EQ(u"One Microsoft Way", builder->InsertField(u" USERADDRESS ")->get_Result());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"FieldOptions.FieldOptionsCurrentUser.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"FieldOptions.FieldOptionsCurrentUser.docx");

        ASSERT_TRUE(doc->get_FieldOptions()->get_CurrentUser() == nullptr);

        auto fieldUserName = System::DynamicCast<FieldUserName>(doc->get_Range()->get_Fields()->idx_get(0));

        ASSERT_TRUE(fieldUserName->get_UserName() == nullptr);
        ASSERT_EQ(u"Default User", fieldUserName->get_Result());

        auto fieldUserInitials = System::DynamicCast<FieldUserInitials>(doc->get_Range()->get_Fields()->idx_get(1));

        ASSERT_TRUE(fieldUserInitials->get_UserInitials() == nullptr);
        ASSERT_EQ(u"D. U.", fieldUserInitials->get_Result());

        auto fieldUserAddress = System::DynamicCast<FieldUserAddress>(doc->get_Range()->get_Fields()->idx_get(2));

        ASSERT_TRUE(fieldUserAddress->get_UserAddress() == nullptr);
        ASSERT_EQ(u"One Microsoft Way", fieldUserAddress->get_Result());
    }

    void FieldOptionsFileName()
    {
        //ExStart
        //ExFor:FieldOptions.FileName
        //ExFor:FieldFileName
        //ExFor:FieldFileName.IncludeFullPath
        //ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToDocumentEnd();
        builder->Writeln();

        // This FILENAME field will display the file name of the document we opened
        auto field = System::DynamicCast<FieldFileName>(builder->InsertField(FieldType::FieldFileName, true));
        field->Update();

        ASSERT_EQ(u" FILENAME ", field->GetFieldCode());
        ASSERT_EQ(u"Document.docx", field->get_Result());

        builder->Writeln();

        // By default, the FILENAME field does not show the full path, and we can change this
        field = System::DynamicCast<FieldFileName>(builder->InsertField(FieldType::FieldFileName, true));
        field->set_IncludeFullPath(true);

        // We can override the values displayed by our FILENAME fields by setting this attribute
        doc->get_FieldOptions()->set_FileName(u"FieldOptions.FILENAME.docx");
        field->Update();

        ASSERT_EQ(u" FILENAME  \\p", field->GetFieldCode());
        ASSERT_EQ(u"FieldOptions.FILENAME.docx", field->get_Result());

        doc->UpdateFields();
        doc->Save(ArtifactsDir + doc->get_FieldOptions()->get_FileName());
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"FieldOptions.FILENAME.docx");

        ASSERT_TRUE(doc->get_FieldOptions()->get_FileName() == nullptr);
        TestUtil::VerifyField(FieldType::FieldFileName, u" FILENAME ", u"FieldOptions.FILENAME.docx", doc->get_Range()->get_Fields()->idx_get(0));
    }

    void FieldOptionsBidi()
    {
        //ExStart
        //ExFor:FieldOptions.IsBidiTextSupportedOnUpdate
        //ExSummary:Shows how to use FieldOptions to ensure that bi-directional text is properly supported during the field update.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Ensure that any field operation involving right-to-left text is performed correctly
        doc->get_FieldOptions()->set_IsBidiTextSupportedOnUpdate(true);

        // Use a document builder to insert a field which contains right-to-left text
        SharedPtr<FormField> comboBox =
            builder->InsertComboBox(u"MyComboBox", MakeArray<String>({u"עֶשְׂרִים", u"שְׁלוֹשִׁים", u"אַרְבָּעִים", u"חֲמִשִּׁים", u"שִׁשִּׁים"}), 0);
        comboBox->set_CalculateOnExit(true);

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"FieldOptions.FieldOptionsBidi.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"FieldOptions.FieldOptionsBidi.docx");

        ASSERT_FALSE(doc->get_FieldOptions()->get_IsBidiTextSupportedOnUpdate());

        comboBox = doc->get_Range()->get_FormFields()->idx_get(0);

        ASSERT_EQ(u"עֶשְׂרִים", comboBox->get_Result());
    }

    void FieldOptionsLegacyNumberFormat()
    {
        //ExStart
        //ExFor:FieldOptions.LegacyNumberFormat
        //ExSummary:Shows how use FieldOptions to change the number format.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Field> field = builder->InsertField(u"= 2 + 3 \\# $##");

        ASSERT_EQ(u"$ 5", field->get_Result());

        doc->get_FieldOptions()->set_LegacyNumberFormat(true);
        field->Update();

        ASSERT_EQ(u"$5", field->get_Result());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_FALSE(doc->get_FieldOptions()->get_LegacyNumberFormat());
        TestUtil::VerifyField(FieldType::FieldFormula, u"= 2 + 3 \\# $##", u"$5", doc->get_Range()->get_Fields()->idx_get(0));
    }

    void FieldOptionsPreProcessCulture()
    {
        //ExStart
        //ExFor:FieldOptions.PreProcessCulture
        //ExSummary:Shows how to set the preprocess culture.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        doc->get_FieldOptions()->set_PreProcessCulture(MakeObject<System::Globalization::CultureInfo>(u"de-DE"));

        SharedPtr<Field> field = builder->InsertField(u" DOCPROPERTY CreateTime");

        // Conforming to the German culture, the date/time will be presented in the "dd.mm.yyyy hh:mm" format
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->get_Result(), u"\\d{2}[.]\\d{2}[.]\\d{4} \\d{2}[:]\\d{2}")->get_Success());

        doc->get_FieldOptions()->set_PreProcessCulture(System::Globalization::CultureInfo::get_InvariantCulture());
        field->Update();

        // After switching to the invariant culture, the date/time will be presented in the "mm/dd/yyyy hh:mm" format
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(field->get_Result(), u"\\d{2}[/]\\d{2}[/]\\d{4} \\d{2}[:]\\d{2}")->get_Success());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_TRUE(doc->get_FieldOptions()->get_PreProcessCulture() == nullptr);
        ASSERT_TRUE(System::Text::RegularExpressions::Regex::Match(doc->get_Range()->get_Fields()->idx_get(0)->get_Result(),
                                                                   u"\\d{2}[/]\\d{2}[/]\\d{4} \\d{2}[:]\\d{2}")
                        ->get_Success());
    }

    void FieldOptionsToaCategories()
    {
        //ExStart
        //ExFor:FieldOptions.ToaCategories
        //ExFor:ToaCategories
        //ExFor:ToaCategories.Item(Int32)
        //ExFor:ToaCategories.DefaultCategories
        //ExSummary:Shows how to specify a table of authorities categories for a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // There are default category values we can use, or we can make our own like this
        auto toaCategories = MakeObject<ToaCategories>();
        doc->get_FieldOptions()->set_ToaCategories(toaCategories);

        toaCategories->idx_set(1, u"My Category 1");
        // Replaces default value "Cases"
        toaCategories->idx_set(2, u"My Category 2");
        // Replaces default value "Statutes"

        // Even if we changed the categories in the FieldOptions object, the default categories are still available here
        ASSERT_EQ(u"Cases", ToaCategories::get_DefaultCategories()->idx_get(1));
        ASSERT_EQ(u"Statutes", ToaCategories::get_DefaultCategories()->idx_get(2));

        // Insert 2 tables of authorities, one per category
        builder->InsertField(u"TOA \\c 1 \\h", nullptr);
        builder->InsertField(u"TOA \\c 2 \\h", nullptr);
        builder->InsertBreak(BreakType::PageBreak);

        // Insert TOA entries across 2 categories
        builder->InsertField(u"TA \\c 2 \\l \"entry 1\"");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertField(u"TA \\c 1 \\l \"entry 2\"");
        builder->InsertBreak(BreakType::PageBreak);
        builder->InsertField(u"TA \\c 2 \\l \"entry 3\"");

        doc->UpdateFields();
        doc->Save(ArtifactsDir + u"FieldOptions.TOA.Categories.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"FieldOptions.TOA.Categories.docx");

        ASSERT_TRUE(doc->get_FieldOptions()->get_ToaCategories() == nullptr);

        TestUtil::VerifyField(FieldType::FieldTOA, u"TOA \\c 1 \\h", u"My Category 1\rentry 2\t3\r", doc->get_Range()->get_Fields()->idx_get(0));
        TestUtil::VerifyField(FieldType::FieldTOA, u"TOA \\c 2 \\h", String(u"My Category 2\r") + u"entry 1\t2\r" + u"entry 3\t4\r",
                              doc->get_Range()->get_Fields()->idx_get(1));
    }

    void FieldOptionsUseInvariantCultureNumberFormat()
    {
        //ExStart
        //ExFor:FieldOptions.UseInvariantCultureNumberFormat
        //ExSummary:Shows how to format numbers according to the invariant culture.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"de-DE"));
        SharedPtr<Field> field = builder->InsertField(u" = 1234567,89 \\# $#,###,###.##");
        field->Update();

        // The combination of field, number format and thread culture can sometimes produce an unsuitable result
        ASSERT_FALSE(doc->get_FieldOptions()->get_UseInvariantCultureNumberFormat());
        ASSERT_EQ(u"$1234567,89 .     ", field->get_Result());

        // We can set this attribute to avoid changing the whole thread culture just for numeric formats
        doc->get_FieldOptions()->set_UseInvariantCultureNumberFormat(true);
        field->Update();
        ASSERT_EQ(u"$1.234.567,89", field->get_Result());
        //ExEnd

        doc = DocumentHelper::SaveOpen(doc);

        ASSERT_FALSE(doc->get_FieldOptions()->get_UseInvariantCultureNumberFormat());
        TestUtil::VerifyField(FieldType::FieldFormula, u" = 1234567,89 \\# $#,###,###.##", u"$1.234.567,89", doc->get_Range()->get_Fields()->idx_get(0));
    }

    //ExStart
    //ExFor:FieldOptions.FieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
    //ExSummary:Shows how to specify a culture defining date/time formatting on per field basis.
    void DefineDateTimeFormatting()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(FieldType::FieldTime, true);

        doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);

        // Set a provider that returns a culture object specific for each field
        doc->get_FieldOptions()->set_FieldUpdateCultureProvider(MakeObject<ExFieldOptions::FieldUpdateCultureProvider>());

        auto fieldDate = System::DynamicCast<FieldTime>(doc->get_Range()->get_Fields()->idx_get(0));
        if (fieldDate->get_LocaleId() != (int)EditingLanguage::Russian)
        {
            fieldDate->set_LocaleId((int)EditingLanguage::Russian);
        }

        doc->Save(ArtifactsDir + u"FieldOptions.UpdateDateTimeFormatting.pdf");
    }

private:
    /// <summary>
    /// Provides a CultureInfo object that should be used during the update of a particular field.
    /// </summary>
    class FieldUpdateCultureProvider : public IFieldUpdateCultureProvider
    {
    public:
        /// <summary>
        /// Returns a CultureInfo object to be used during the field's update.
        /// </summary>
        SharedPtr<System::Globalization::CultureInfo> GetCulture(String name, SharedPtr<Field> field) override
        {
            const String& switch_value_0 = name;
            if (switch_value_0 == u"ru-RU")
            {
                auto culture = MakeObject<System::Globalization::CultureInfo>(name, false);
                SharedPtr<System::Globalization::DateTimeFormatInfo> format = culture->get_DateTimeFormat();
                format->set_MonthNames(MakeArray<String>({u"месяц 1", u"месяц 2", u"месяц 3", u"месяц 4", u"месяц 5", u"месяц 6", u"месяц 7",
                                                                          u"месяц 8", u"месяц 9", u"месяц 10", u"месяц 11", u"месяц 12", u""}));
                format->set_MonthGenitiveNames(format->get_MonthNames());
                format->set_AbbreviatedMonthNames(MakeArray<String>(
                    {u"мес 1", u"мес 2", u"мес 3", u"мес 4", u"мес 5", u"мес 6", u"мес 7", u"мес 8", u"мес 9", u"мес 10", u"мес 11", u"мес 12", u""}));
                format->set_AbbreviatedMonthGenitiveNames(format->get_AbbreviatedMonthNames());
                format->set_DayNames(MakeArray<String>(
                    {u"день недели 7", u"день недели 1", u"день недели 2", u"день недели 3", u"день недели 4", u"день недели 5", u"день недели 6"}));
                format->set_AbbreviatedDayNames(
                    MakeArray<String>({u"день 7", u"день 1", u"день 2", u"день 3", u"день 4", u"день 5", u"день 6"}));
                format->set_ShortestDayNames(MakeArray<String>({u"д7", u"д1", u"д2", u"д3", u"д4", u"д5", u"д6"}));
                format->set_AMDesignator(u"До полудня");
                format->set_PMDesignator(u"После полудня");
                const String pattern = u"yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                format->set_LongDatePattern(pattern);
                format->set_LongTimePattern(pattern);
                format->set_ShortDatePattern(pattern);
                format->set_ShortTimePattern(pattern);
                return culture;
            }
            else if (switch_value_0 == u"en-US")
            {
                return MakeObject<System::Globalization::CultureInfo>(name, false);
            }
            else if (true)
            {
                return nullptr;
            }
        }
    };
    //ExEnd
};

} // namespace ApiExamples
