#pragma once

#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/EditableRange.h>
#include <Aspose.Words.Cpp/EditableRangeEnd.h>
#include <Aspose.Words.Cpp/EditableRangeStart.h>
#include <Aspose.Words.Cpp/ProtectionType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Settings/WriteProtection.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents { namespace Protect_or_Encrypt_Document {

class DocumentProtection : public DocsExamplesBase
{
public:
    void PasswordProtection()
    {
        //ExStart:PasswordProtection
        auto doc = MakeObject<Document>();

        // Apply document protection.
        doc->Protect(ProtectionType::NoProtection, u"password");

        doc->Save(ArtifactsDir + u"DocumentProtection.PasswordProtection.docx");
        //ExEnd:PasswordProtection
    }

    void AllowOnlyFormFieldsProtect()
    {
        //ExStart:AllowOnlyFormFieldsProtect
        // Insert two sections with some text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Text added to a document.");

        // A document protection only works when document protection is turned and only editing in form fields is allowed.
        doc->Protect(ProtectionType::AllowOnlyFormFields, u"password");

        // Save the protected document.
        doc->Save(ArtifactsDir + u"DocumentProtection.AllowOnlyFormFieldsProtect.docx");
        //ExEnd:AllowOnlyFormFieldsProtect
    }

    void RemoveDocumentProtection()
    {
        //ExStart:RemoveDocumentProtection
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Text added to a document.");

        // Documents can have protection removed either with no password, or with the correct password.
        doc->Unprotect();
        doc->Protect(ProtectionType::ReadOnly, u"newPassword");
        doc->Unprotect(u"newPassword");

        doc->Save(ArtifactsDir + u"DocumentProtection.RemoveDocumentProtection.docx");
        //ExEnd:RemoveDocumentProtection
    }

    void UnrestrictedEditableRegions()
    {
        //ExStart:UnrestrictedEditableRegions
        // Upload a document and make it as read-only.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        auto builder = MakeObject<DocumentBuilder>(doc);

        doc->Protect(ProtectionType::ReadOnly, u"MyPassword");

        builder->Writeln(String(u"Hello world! Since we have set the document's protection level to read-only, ") +
                         u"we cannot edit this paragraph without the password.");

        // Start an editable range.
        SharedPtr<EditableRangeStart> edRangeStart = builder->StartEditableRange();
        // An EditableRange object is created for the EditableRangeStart that we just made.
        SharedPtr<EditableRange> editableRange = edRangeStart->get_EditableRange();

        // Put something inside the editable range.
        builder->Writeln(u"Paragraph inside first editable range");

        // An editable range is well-formed if it has a start and an end.
        SharedPtr<EditableRangeEnd> edRangeEnd = builder->EndEditableRange();

        builder->Writeln(u"This paragraph is outside any editable ranges, and cannot be edited.");

        doc->Save(ArtifactsDir + u"DocumentProtection.UnrestrictedEditableRegions.docx");
        //ExEnd:UnrestrictedEditableRegions
    }

    void UnrestrictedSection()
    {
        //ExStart:UnrestrictedSection
        // Insert two sections with some text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Section 1. Unprotected.");
        builder->InsertBreak(BreakType::SectionBreakContinuous);
        builder->Writeln(u"Section 2. Protected.");

        // Section protection only works when document protection is turned and only editing in form fields is allowed.
        doc->Protect(ProtectionType::AllowOnlyFormFields, u"password");

        // By default, all sections are protected, but we can selectively turn protection off.
        doc->get_Sections()->idx_get(0)->set_ProtectedForForms(false);
        doc->Save(ArtifactsDir + u"DocumentProtection.UnrestrictedSection.docx");

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentProtection.UnrestrictedSection.docx");
        ASSERT_FALSE(doc->get_Sections()->idx_get(0)->get_ProtectedForForms());
        ASSERT_TRUE(doc->get_Sections()->idx_get(1)->get_ProtectedForForms());
        //ExEnd:UnrestrictedSection
    }

    void GetProtectionType()
    {
        //ExStart:GetProtectionType
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        ProtectionType protectionType = doc->get_ProtectionType();
        //ExEnd:GetProtectionType
    }

    void ReadOnlyProtection()
    {
        //ExStart:ReadOnlyProtection
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Open document as read-only");

        // Enter a password that's up to 15 characters long.
        doc->get_WriteProtection()->SetPassword(u"MyPassword");

        // Make the document as read-only.
        doc->get_WriteProtection()->set_ReadOnlyRecommended(true);

        // Apply write protection as read-only.
        doc->Protect(ProtectionType::ReadOnly);
        doc->Save(ArtifactsDir + u"DocumentProtection.ReadOnlyProtection.docx");
        //ExEnd:ReadOnlyProtection
    }

    void RemoveReadOnlyRestriction()
    {
        //ExStart:RemoveReadOnlyRestriction
        auto doc = MakeObject<Document>();

        // Enter a password that's up to 15 characters long.
        doc->get_WriteProtection()->SetPassword(u"MyPassword");

        // Remove the read-only option.
        doc->get_WriteProtection()->set_ReadOnlyRecommended(false);

        // Apply write protection without any protection.
        doc->Protect(ProtectionType::NoProtection);
        doc->Save(ArtifactsDir + u"DocumentProtection.RemoveReadOnlyRestriction.docx");
        //ExEnd:RemoveReadOnlyRestriction
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Protect_or_Encrypt_Document
