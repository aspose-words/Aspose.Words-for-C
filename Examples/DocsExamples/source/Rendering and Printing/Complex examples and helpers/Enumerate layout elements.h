#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEntityType.h>
#include <Aspose.Words.Cpp/Layout/Public/LayoutEnumerator.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <drawing/graphics.h>
#include <drawing/pen.h>
#include <system/enum_helpers.h>

#include "DocsExamplesBase.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Layout;

namespace DocsExamples { namespace Complex_examples_and_helpers {

class EnumerateLayoutElements : public DocsExamplesBase
{
public:
    void GetLayoutElements();
};

class LayoutInfoWriter : public System::Object
{
public:
    static void Run(System::SharedPtr<LayoutEnumerator> layoutEnumerator);

private:
    /// <summary>
    /// Enumerates forward through each layout element in the document and prints out details of each element.
    /// </summary>
    static void DisplayLayoutElements(System::SharedPtr<LayoutEnumerator> layoutEnumerator, System::String padding);
    /// <summary>
    /// Displays information about the current layout entity to the console.
    /// </summary>
    static void DisplayEntityInfo(System::SharedPtr<LayoutEnumerator> layoutEnumerator, System::String padding);
    /// <summary>
    /// Returns a string of spaces for padding purposes.
    /// </summary>
    static System::String AddPadding(System::String padding);
};

class OutlineLayoutEntitiesRenderer : public System::Object
{
public:
    static void Run(System::SharedPtr<Document> doc, System::SharedPtr<LayoutEnumerator> layoutEnumerator, System::String folderPath);

private:
    /// <summary>
    /// Adds a colored border around each layout element on the page.
    /// </summary>
    static void AddBoundingBoxToElementsOnPage(System::SharedPtr<LayoutEnumerator> layoutEnumerator, System::SharedPtr<System::Drawing::Graphics> g);
    /// <summary>
    /// Returns a different colored pen for each entity type.
    /// </summary>
    static System::SharedPtr<System::Drawing::Pen> GetColoredPenFromType(LayoutEntityType type);
    /// <summary>
    /// Converts a value in points to pixels.
    /// </summary>
    static int PointToPixel(float value, double resolution);
};

}} // namespace DocsExamples::Complex_examples_and_helpers
