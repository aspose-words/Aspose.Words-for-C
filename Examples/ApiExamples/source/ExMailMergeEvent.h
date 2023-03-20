#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldMergeField.h>
#include <Aspose.Words.Cpp/MailMerging/FieldMergingArgs.h>
#include <Aspose.Words.Cpp/MailMerging/IFieldMergingCallback.h>
#include <Aspose.Words.Cpp/MailMerging/ImageFieldMergingArgs.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/exceptions.h>
#include <system/object_ext.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::MailMerging;

namespace ApiExamples {

class ExMailMergeEvent : public ApiExampleBase
{
public:
    //ExStart
    //ExFor:DocumentBuilder.InsertHtml(String)
    //ExFor:MailMerge.FieldMergingCallback
    //ExFor:IFieldMergingCallback
    //ExFor:FieldMergingArgs
    //ExFor:FieldMergingArgsBase
    //ExFor:FieldMergingArgsBase.Field
    //ExFor:FieldMergingArgsBase.DocumentFieldName
    //ExFor:FieldMergingArgsBase.Document
    //ExFor:IFieldMergingCallback.FieldMerging
    //ExFor:FieldMergingArgs.Text
    //ExSummary:Shows how to execute a mail merge with a custom callback that handles merge data in the form of HTML documents.
    void MergeHtml()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD  html_Title  \\b Content");
        builder->InsertField(u"MERGEFIELD  html_Body  \\b Content");

        ArrayPtr<SharedPtr<System::Object>> mergeData = MakeArray<SharedPtr<System::Object>>(
            {System::ObjectExt::Box<String>(String(u"<html>") + u"<h1>" + u"<span style=\"color: #0000ff; font-family: Arial;\">Hello World!</span>" +
                                            u"</h1>" + u"</html>"),
             System::ObjectExt::Box<String>(
                 String(u"<html>") + u"<blockquote>" +
                 u"<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>" +
                 u"</blockquote>" + u"</html>")});

        doc->get_MailMerge()->set_FieldMergingCallback(MakeObject<ExMailMergeEvent::HandleMergeFieldInsertHtml>());
        doc->get_MailMerge()->Execute(MakeArray<String>({u"html_Title", u"html_Body"}), mergeData);

        doc->Save(ArtifactsDir + u"MailMergeEvent.MergeHtml.docx");
    }

    /// <summary>
    /// If the mail merge encounters a MERGEFIELD whose name starts with the "html_" prefix,
    /// this callback parses its merge data as HTML content and adds the result to the document location of the MERGEFIELD.
    /// </summary>
    class HandleMergeFieldInsertHtml : public IFieldMergingCallback
    {
    private:
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        void FieldMerging(SharedPtr<FieldMergingArgs> args) override
        {
            if (args->get_DocumentFieldName().StartsWith(u"html_") && args->get_Field()->GetFieldCode().Contains(u"\\b"))
            {
                // Add parsed HTML data to the document's body.
                auto builder = MakeObject<DocumentBuilder>(args->get_Document());
                builder->MoveToMergeField(args->get_DocumentFieldName());
                builder->InsertHtml(System::ObjectExt::Unbox<String>(args->get_FieldValue()));

                // Since we have already inserted the merged content manually,
                // we will not need to respond to this event by returning content via the "Text" property.
                args->set_Text(String::Empty);
            }
        }

        void ImageFieldMerging(SharedPtr<ImageFieldMergingArgs> args) override
        {
            // Do nothing.
        }
    };
    //ExEnd

};

} // namespace ApiExamples
