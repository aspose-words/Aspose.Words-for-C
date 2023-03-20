#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Rendering/PageInfo.h>
#include <drawing/bitmap.h>
#include <drawing/color.h>
#include <drawing/graphics.h>
#include <drawing/graphics_unit.h>
#include <drawing/pen.h>
#include <drawing/pens.h>
#include <drawing/size.h>
#include <drawing/size_f.h>
#include <drawing/solid_brush.h>
#include <drawing/text/text_rendering_hint.h>
#include <system/details/dispose_guard.h>
#include <system/math.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Rendering;

namespace ApiExamples {

class ExRendering : public ApiExampleBase
{

public:

    void Thumbnails()
    {
        //ExStart
        //ExFor:Document.RenderToScale
        //ExSummary:Shows how to the individual pages of a document to graphics to create one image with thumbnails of all pages.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Calculate the number of rows and columns that we will fill with thumbnails.
        const int thumbColumns = 2;
        int remainder;
        int thumbRows = System::Math::DivRem(doc->get_PageCount(), thumbColumns, remainder);

        if (remainder > 0)
        {
            thumbRows++;
        }

        // Scale the thumbnails relative to the size of the first page.
        const float scale = 0.25f;
        System::Drawing::Size thumbSize = doc->GetPageInfo(0)->GetSizeInPixels(scale, 96.0f);

        // Calculate the size of the image that will contain all the thumbnails.
        int imgWidth = thumbSize.get_Width() * thumbColumns;
        int imgHeight = thumbSize.get_Height() * thumbRows;

        {
            auto img = MakeObject<System::Drawing::Bitmap>(imgWidth, imgHeight);
            {
                SharedPtr<System::Drawing::Graphics> gr = System::Drawing::Graphics::FromImage(img);
                gr->set_TextRenderingHint(System::Drawing::Text::TextRenderingHint::AntiAliasGridFit);

                // Fill the background, which is transparent by default, in white.
                gr->FillRectangle(MakeObject<System::Drawing::SolidBrush>(System::Drawing::Color::get_White()), 0, 0, imgWidth, imgHeight);

                for (int pageIndex = 0; pageIndex < doc->get_PageCount(); pageIndex++)
                {
                    int columnIdx;
                    int rowIdx = System::Math::DivRem(pageIndex, thumbColumns, columnIdx);

                    // Specify where we want the thumbnail to appear.
                    float thumbLeft = static_cast<float>(columnIdx * thumbSize.get_Width());
                    float thumbTop = static_cast<float>(rowIdx * thumbSize.get_Height());

                    // Render a page as a thumbnail, and then frame it in a rectangle of the same size.
                    System::Drawing::SizeF size = doc->RenderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);
                    gr->DrawRectangle(System::Drawing::Pens::get_Black(), thumbLeft, thumbTop, size.get_Width(), size.get_Height());
                }

                img->Save(ArtifactsDir + u"Rendering.Thumbnails.png");
            }
        }
        //ExEnd
    }
};

} // namespace ApiExamples
